//! Bounded, source-borrowing wire seams for the Keynote media lifecycle graph.
//!
//! The native `KN.SlideArchive`, `KN.BuildArchive`, and
//! `KN.BuildChunkArchive` messages are larger than the small graph needed by a
//! media duplicate/remove transaction.  This module deliberately projects only
//! the reference and UUID edges used by that transaction.  A strict handwritten
//! pass still walks every field (including unknown groups), while a private
//! Buffa lazy view cross-checks the selected fields without materialising a
//! generated repeated collection.  The input bytes remain the preservation
//! authority: rewrites copy every untouched field span verbatim and replace
//! only selected canonical scalar edges.
//!
//! No package identity, object archive, UUID registry, movie title/caption
//! graph, or data-map ownership rule belongs here.  The Keynote package owner
//! supplies those witnesses and publishes a candidate only after its package
//! transaction has validated them.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict pass, lazy cross-check, and raw-preserving writer are kept together."
)]

use core::{fmt, mem::size_of};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_media_lifecycle_generated::LitchiIwaProjection as projection;

const MAX_RECURSION_LIMIT: u32 = 64;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

const SLIDE_STYLE_FIELD: u32 = 1;
const SLIDE_BUILDS_FIELD: u32 = 2;
const SLIDE_DEPRECATED_BUILD_CHUNKS_FIELD: u32 = 3;
const SLIDE_TRANSITION_FIELD: u32 = 4;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_IN_DOCUMENT_FIELD: u32 = 19;
const SLIDE_DRAWABLES_Z_ORDER_FIELD: u32 = 42;
const SLIDE_BUILD_CHUNKS_FIELD: u32 = 43;
const TRANSITION_ATTRIBUTES_FIELD: u32 = 2;

const BUILD_DRAWABLE_FIELD: u32 = 1;
const BUILD_DELIVERY_FIELD: u32 = 2;
const BUILD_ATTRIBUTES_FIELD: u32 = 4;

const CHUNK_BUILD_FIELD: u32 = 1;
const CHUNK_AUTOMATIC_FIELD: u32 = 5;
const CHUNK_REFERENT_FIELD: u32 = 6;
const CHUNK_IDENTIFIER_FIELD: u32 = 7;
const CHUNK_BUILD_ID_FIELD: u32 = 8;
const CHUNK_IDENTIFIER_UUID_FIELD: u32 = 1;
const CHUNK_IDENTIFIER_INDEX_FIELD: u32 = 2;

const UUID_LOWER_FIELD: u32 = 1;
const UUID_UPPER_FIELD: u32 = 2;

const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;

/// Finite source, output, field, work, topology, and nesting limits for one
/// media-lifecycle operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_references: usize,
    max_depth: u32,
}

impl DecodeOptions {
    /// Construct an explicit finite policy.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_references: usize,
        max_depth: u32,
    ) -> Self {
        Self {
            max_message_bytes,
            max_output_bytes,
            max_fields,
            max_work_bytes,
            max_references,
            max_depth,
        }
    }

    /// Construct a conservative operation-local policy from one borrowed
    /// source payload.  Callers handling untrusted package data should tighten
    /// these ceilings to their package semantic limits.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(2).max(1),
            bytes.saturating_mul(8).max(1),
            bytes.saturating_mul(32).max(1),
            bytes.saturating_mul(2).max(1),
            16,
        )
    }

    /// Maximum source message bytes.
    #[must_use]
    pub const fn max_message_bytes(self) -> usize {
        self.max_message_bytes
    }

    /// Maximum candidate output bytes.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Maximum complete strict field visits.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Maximum aggregate decoder and rewrite work bytes.
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    /// Maximum references retained by one semantic list.
    #[must_use]
    pub const fn max_references(self) -> usize {
        self.max_references
    }

    /// Maximum supported protobuf nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Set a source-byte ceiling.
    #[must_use]
    pub const fn with_max_message_bytes(mut self, maximum: usize) -> Self {
        self.max_message_bytes = maximum;
        self
    }

    /// Set a candidate-output ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Set a strict field ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Set an aggregate work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Set a per-list reference ceiling.
    #[must_use]
    pub const fn with_max_references(mut self, maximum: usize) -> Self {
        self.max_references = maximum;
        self
    }

    /// Set the nesting ceiling.
    #[must_use]
    pub const fn with_max_depth(mut self, maximum: u32) -> Self {
        self.max_depth = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            // Unknown bytes stay source-owned by the handwritten pass, while
            // this finite allowance keeps the lazy parity view bounded.
            .with_unknown_field_limit(self.max_fields)
            // The sidecar has no repeated owned fields.
            .with_element_memory_limit(0)
            .with_recursion_limit(self.max_depth)
    }
}

/// A bounded resource classification returned by [`DecodeError`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Input or nested message bytes exceeded the source ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Strict field visits exceeded the field ceiling.
    Fields { observed: usize, maximum: usize },
    /// Aggregate parser, Buffa, and rewrite work exceeded the work ceiling.
    Work { observed: usize, maximum: usize },
    /// Candidate output exceeded the output ceiling.
    OutputBytes { observed: usize, maximum: usize },
    /// A semantic repeated reference list exceeded its topology ceiling.
    References { observed: usize, maximum: usize },
    /// Configured or observed nesting exceeded the finite ceiling.
    Nesting { observed: u32, maximum: u32 },
}

/// A strict Keynote lifecycle decode or rewrite failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Resource(DecodeLimit),
    Missing(&'static str),
    Duplicate(&'static str),
    WrongWire(&'static str),
    NonCanonical(&'static str),
    Unsupported(&'static str),
    Projection,
    Allocation(&'static str),
}

impl DecodeError {
    const fn invalid() -> Self {
        Self {
            kind: DecodeErrorKind::Unsupported("invalid Keynote lifecycle wire"),
        }
    }

    const fn missing(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Missing(field),
        }
    }

    const fn duplicate(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Duplicate(field),
        }
    }

    const fn wrong_wire(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::WrongWire(field),
        }
    }

    const fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason),
        }
    }

    const fn unsupported(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Unsupported(reason),
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    const fn allocation(resource: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation(resource),
        }
    }

    const fn limited(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Resource(limit),
        }
    }

    /// Return the bounded-resource classification, when applicable.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match &self.kind {
            DecodeErrorKind::Resource(limit) => Some(*limit),
            _ => None,
        }
    }

    /// Return the duplicated known field, when applicable.
    #[must_use]
    pub const fn duplicate_field(&self) -> Option<&'static str> {
        match &self.kind {
            DecodeErrorKind::Duplicate(field) => Some(*field),
            _ => None,
        }
    }

    /// Return the unsupported topology reason, when applicable.
    #[must_use]
    pub const fn unsupported_reason(&self) -> Option<&'static str> {
        match &self.kind {
            DecodeErrorKind::Unsupported(reason) => Some(*reason),
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Resource(DecodeLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "Keynote lifecycle wire bytes exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "Keynote lifecycle wire fields exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "Keynote lifecycle wire work exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::OutputBytes { observed, maximum }) => write!(
                formatter,
                "Keynote lifecycle output bytes exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::References { observed, maximum }) => write!(
                formatter,
                "Keynote lifecycle references exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Keynote lifecycle nesting exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Missing(field) => write!(formatter, "missing required field {field}"),
            DecodeErrorKind::Duplicate(field) => write!(formatter, "duplicate field {field}"),
            DecodeErrorKind::WrongWire(field) => write!(formatter, "wrong wire type for {field}"),
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::Unsupported(reason) => {
                write!(formatter, "unsupported topology: {reason}")
            },
            DecodeErrorKind::Projection => formatter.write_str(
                "Keynote lifecycle strict wire pass disagrees with the Buffa lazy projection",
            ),
            DecodeErrorKind::Allocation(resource) => {
                write!(formatter, "allocation failed for {resource}")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self {
            kind: DecodeErrorKind::Wire(error),
        }
    }
}

/// Exact resource consumption for a successful decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl DecodeReport {
    /// Source bytes inspected by the operation.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Strict fields visited, including unknown fields and group contents.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate bounded work charged by the operation.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Greatest nesting depth observed.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Number of fallible owned allocations performed by the handwritten pass.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source bytes retained by borrowed snapshots.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Temporary scratch bytes charged by this operation.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Exact 128-bit UUID scalar pair used by a build chunk.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Uuid {
    lower: u64,
    upper: u64,
}

impl Uuid {
    /// Construct a UUID from its native scalar halves.
    #[must_use]
    pub const fn new(lower: u64, upper: u64) -> Self {
        Self { lower, upper }
    }

    /// Lower 64-bit half.
    #[must_use]
    pub const fn lower(self) -> u64 {
        self.lower
    }

    /// Upper 64-bit half.
    #[must_use]
    pub const fn upper(self) -> u64 {
        self.upper
    }
}

/// A borrowed native `TSP.Reference` edge.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Reference<'source> {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
    raw: &'source [u8],
}

impl<'source> Reference<'source> {
    /// Native object identifier.
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }

    /// Deprecated native type scalar, when present.
    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }

    /// Deprecated external-reference flag, when present.
    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }

    /// Exact source payload of this reference, excluding its enclosing field.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

/// A borrowed UUID edge.  The exact source payload is retained for callers
/// that need to preserve an untouched nested span.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct UuidSnapshot<'source> {
    uuid: Uuid,
    raw: &'source [u8],
}

impl<'source> UuidSnapshot<'source> {
    /// Parsed UUID halves.
    #[must_use]
    pub const fn uuid(self) -> Uuid {
        self.uuid
    }

    /// Exact source UUID payload.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

/// Borrowed scalar facts from one complete `KN.SlideArchive` payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideLifecycleSnapshot<'source> {
    source: &'source [u8],
    style: Reference<'source>,
    in_document: bool,
    owned_drawables: Vec<Reference<'source>>,
    drawables_z_order: Vec<Reference<'source>>,
    builds: Vec<Reference<'source>>,
    build_chunks: Vec<Reference<'source>>,
}

impl<'source> SlideLifecycleSnapshot<'source> {
    /// Exact validated source payload.
    #[must_use]
    pub const fn source(&self) -> &'source [u8] {
        self.source
    }

    /// Slide style reference.
    #[must_use]
    pub const fn style(&self) -> Reference<'source> {
        self.style
    }

    /// Native `inDocument` value.
    #[must_use]
    pub const fn in_document(&self) -> bool {
        self.in_document
    }

    /// Owned drawables in source order.
    #[must_use]
    pub fn owned_drawables(&self) -> impl ExactSizeIterator<Item = Reference<'source>> + '_ {
        self.owned_drawables.iter().copied()
    }

    /// Drawable z-order references in source order.
    #[must_use]
    pub fn drawables_z_order(&self) -> impl ExactSizeIterator<Item = Reference<'source>> + '_ {
        self.drawables_z_order.iter().copied()
    }

    /// Build object references in source order.
    #[must_use]
    pub fn builds(&self) -> impl ExactSizeIterator<Item = Reference<'source>> + '_ {
        self.builds.iter().copied()
    }

    /// Build-chunk object references in source order.
    #[must_use]
    pub fn build_chunks(&self) -> impl ExactSizeIterator<Item = Reference<'source>> + '_ {
        self.build_chunks.iter().copied()
    }
}

/// Borrowed scalar facts from one `KN.BuildArchive` payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BuildLifecycleSnapshot<'source> {
    source: &'source [u8],
    drawable: Option<Reference<'source>>,
}

impl<'source> BuildLifecycleSnapshot<'source> {
    /// Exact validated source payload.
    #[must_use]
    pub const fn source(self) -> &'source [u8] {
        self.source
    }

    /// Drawable reference, when the build has one.
    #[must_use]
    pub const fn drawable(self) -> Option<Reference<'source>> {
        self.drawable
    }
}

/// Borrowed scalar facts from one `KN.BuildChunkArchive` payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BuildChunkLifecycleSnapshot<'source> {
    source: &'source [u8],
    build: Reference<'source>,
    chunk_identifier: Option<UuidSnapshot<'source>>,
    build_id: Option<UuidSnapshot<'source>>,
}

impl<'source> BuildChunkLifecycleSnapshot<'source> {
    /// Exact validated source payload.
    #[must_use]
    pub const fn source(self) -> &'source [u8] {
        self.source
    }

    /// Referenced build object.
    #[must_use]
    pub const fn build(self) -> Reference<'source> {
        self.build
    }

    /// Nested `buildChunkIdentifier` UUID, when present.
    #[must_use]
    pub const fn chunk_identifier(self) -> Option<UuidSnapshot<'source>> {
        self.chunk_identifier
    }

    /// Direct `buildId` UUID, when present.
    #[must_use]
    pub const fn build_id(self) -> Option<UuidSnapshot<'source>> {
        self.build_id
    }
}

/// One checked object-identifier rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct IdentifierRewrite {
    source: u64,
    target: u64,
}

impl IdentifierRewrite {
    /// Construct a source-to-target rewrite request.  Zero values are
    /// rejected when the request is prepared against a source payload.
    #[must_use]
    pub const fn new(source: u64, target: u64) -> Self {
        Self { source, target }
    }

    /// Source identifier.
    #[must_use]
    pub const fn source(self) -> u64 {
        self.source
    }

    /// Target identifier.
    #[must_use]
    pub const fn target(self) -> u64 {
        self.target
    }
}

/// One checked UUID rewrite request.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct UuidRewrite {
    source: Uuid,
    target: Uuid,
}

impl UuidRewrite {
    /// Construct a source-to-target UUID rewrite request.
    #[must_use]
    pub const fn new(source: Uuid, target: Uuid) -> Self {
        Self { source, target }
    }

    /// Source UUID.
    #[must_use]
    pub const fn source(self) -> Uuid {
        self.source
    }

    /// Target UUID.
    #[must_use]
    pub const fn target(self) -> Uuid {
        self.target
    }
}

/// Selector-free, bounded slide-list edit.  The Keynote owner resolves the
/// public slide selector and all package-wide object witnesses before passing
/// these compact wire edits.
#[derive(Debug, Clone, Copy)]
pub struct SlideLifecycleEdit<'edit> {
    remove_identifiers: &'edit [u64],
    remap_identifiers: &'edit [IdentifierRewrite],
    append_owned_drawables: &'edit [u64],
    append_drawables_z_order: &'edit [u64],
    append_builds: &'edit [u64],
    append_build_chunks: &'edit [u64],
}

impl<'edit> SlideLifecycleEdit<'edit> {
    /// Construct an empty edit.
    #[must_use]
    pub const fn empty() -> Self {
        Self {
            remove_identifiers: &[],
            remap_identifiers: &[],
            append_owned_drawables: &[],
            append_drawables_z_order: &[],
            append_builds: &[],
            append_build_chunks: &[],
        }
    }

    /// Set identifiers whose reference fields are removed from every selected
    /// slide list.
    #[must_use]
    pub const fn with_removed_identifiers(mut self, identifiers: &'edit [u64]) -> Self {
        self.remove_identifiers = identifiers;
        self
    }

    /// Set selected object-identifier rewrites applied to every selected slide
    /// reference list.
    #[must_use]
    pub const fn with_identifier_remaps(mut self, remaps: &'edit [IdentifierRewrite]) -> Self {
        self.remap_identifiers = remaps;
        self
    }

    /// Append owned drawable identifiers in native source order.
    #[must_use]
    pub const fn with_owned_drawables(mut self, identifiers: &'edit [u64]) -> Self {
        self.append_owned_drawables = identifiers;
        self
    }

    /// Append drawable z-order identifiers.
    #[must_use]
    pub const fn with_drawables_z_order(mut self, identifiers: &'edit [u64]) -> Self {
        self.append_drawables_z_order = identifiers;
        self
    }

    /// Append build object identifiers.
    #[must_use]
    pub const fn with_builds(mut self, identifiers: &'edit [u64]) -> Self {
        self.append_builds = identifiers;
        self
    }

    /// Append build-chunk object identifiers.
    #[must_use]
    pub const fn with_build_chunks(mut self, identifiers: &'edit [u64]) -> Self {
        self.append_build_chunks = identifiers;
        self
    }
}

/// Build payload rewrite request.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BuildLifecycleEdit {
    drawable: IdentifierRewrite,
}

impl BuildLifecycleEdit {
    /// Rewrite one build's drawable edge.
    #[must_use]
    pub const fn drawable(rewrite: IdentifierRewrite) -> Self {
        Self { drawable: rewrite }
    }
}

/// Build-chunk payload rewrite request.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BuildChunkLifecycleEdit {
    build: Option<IdentifierRewrite>,
    uuid: Option<UuidRewrite>,
}

impl BuildChunkLifecycleEdit {
    /// Construct an edit with no selected edges.
    #[must_use]
    pub const fn empty() -> Self {
        Self {
            build: None,
            uuid: None,
        }
    }

    /// Rewrite the referenced build object.
    #[must_use]
    pub const fn with_build(mut self, rewrite: IdentifierRewrite) -> Self {
        self.build = Some(rewrite);
        self
    }

    /// Rewrite both native UUID edges when present.
    #[must_use]
    pub const fn with_uuid(mut self, rewrite: UuidRewrite) -> Self {
        self.uuid = Some(rewrite);
        self
    }
}

/// Exact accounting for one successful source-preserving rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    changed: bool,
}

impl RewriteReport {
    /// Source bytes inspected.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Candidate output bytes.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Complete strict field visits charged by preparation and readback.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Complete bounded work charged by the rewrite.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Greatest observed nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Number of output allocations.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source and candidate bytes retained by the transaction result.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Temporary scratch bytes charged by the transaction.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Whether any selected edge or list changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A prepared slide rewrite holds an immutable source witness until the
/// caller chooses to publish the candidate.  Preparation performs all source
/// validation, topology checks, output sizing, and aggregate budget checks.
#[derive(Debug)]
pub struct PreparedSlideLifecycleRewrite<'source, 'edit> {
    source: &'source [u8],
    edit: SlideLifecycleEdit<'edit>,
    options: DecodeOptions,
    snapshot: SlideLifecycleSnapshot<'source>,
    output_bytes: usize,
    estimated_work_bytes: usize,
}

impl<'source, 'edit> PreparedSlideLifecycleRewrite<'source, 'edit> {
    /// Return the exact immutable source witness used during preparation.
    #[must_use]
    pub const fn source(&self) -> &'source [u8] {
        self.source
    }

    /// Return the exact candidate size calculated before allocation.
    #[must_use]
    pub const fn output_bytes(&self) -> usize {
        self.output_bytes
    }

    /// Return the source snapshot used to validate the edit.
    #[must_use]
    pub const fn snapshot(&self) -> &SlideLifecycleSnapshot<'source> {
        &self.snapshot
    }

    /// Emit, strictly read back, and return the candidate bytes and report.
    pub fn commit(self) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
        let mut output = Vec::new();
        output
            .try_reserve_exact(self.output_bytes)
            .map_err(|_| DecodeError::allocation("slide lifecycle output"))?;
        let mut budget = Budget::new(self.source, self.options);
        emit_slide(
            self.source,
            &self.snapshot,
            self.edit,
            self.options,
            &mut budget,
            &mut output,
        )?;
        if output.len() != self.output_bytes {
            return Err(DecodeError::projection());
        }
        let readback_options = self
            .options
            .with_max_message_bytes(self.options.max_message_bytes.max(output.len()));
        let (readback, readback_report) =
            decode_slide_lifecycle_with_report(&output, readback_options)?;
        if !same_slide_semantics(&readback, &self.snapshot, self.edit, &mut budget)? {
            return Err(DecodeError::projection());
        }
        let changed = output.as_slice() != self.source;
        let work_bytes = self
            .estimated_work_bytes
            .saturating_add(budget.work_bytes)
            .saturating_add(readback_report.work_bytes);
        Ok((
            output,
            RewriteReport {
                input_bytes: self.source.len(),
                output_bytes: self.output_bytes,
                fields: budget.fields.saturating_add(readback_report.fields),
                work_bytes,
                max_depth: budget.max_depth.max(readback_report.max_depth),
                allocations: 1,
                retained_bytes: self.source.len().saturating_add(self.output_bytes),
                scratch_bytes: budget
                    .scratch_bytes
                    .saturating_add(readback_report.scratch_bytes),
                changed,
            },
        ))
    }
}

/// Decode one complete slide lifecycle payload.
pub fn decode_slide_lifecycle<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<SlideLifecycleSnapshot<'source>, DecodeError> {
    decode_slide_lifecycle_with_report(source, options).map(|(snapshot, _)| snapshot)
}

/// Decode one complete slide lifecycle payload and return exact finite usage.
pub fn decode_slide_lifecycle_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(SlideLifecycleSnapshot<'source>, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = parse_slide(source, options, &mut budget)?;
    // The sidecar intentionally has no repeated slide fields. Repeated
    // ownership/z-order/build envelopes are streamed by the handwritten pass;
    // each selected nested reference is cross-checked through the tiny lazy
    // Buffa scalar view while it is encountered.
    budget.charge_work(source.len())?;
    for reference in snapshot
        .owned_drawables
        .iter()
        .chain(snapshot.drawables_z_order.iter())
        .chain(snapshot.builds.iter())
        .chain(snapshot.build_chunks.iter())
        .chain(std::iter::once(&snapshot.style))
    {
        budget.charge_work(reference.raw.len())?;
        force_reference_projection(*reference, options)?;
    }
    Ok((snapshot, budget.report(source.len())))
}

/// Decode one `KN.BuildArchive` payload.
pub fn decode_build<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<BuildLifecycleSnapshot<'source>, DecodeError> {
    decode_build_with_report(source, options).map(|(snapshot, _)| snapshot)
}

/// Decode one build payload and return exact finite usage.
pub fn decode_build_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(BuildLifecycleSnapshot<'source>, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = parse_build(source, options, &mut budget)?;
    let view: projection::BuildArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    budget.charge_work(source.len())?;
    let projected = view.drawable.get().map_err(DecodeError::from)?;
    if projected.is_some() != snapshot.drawable.is_some() {
        return Err(DecodeError::projection());
    }
    if let (Some(projected), Some(expected)) = (projected.as_ref(), snapshot.drawable) {
        budget.charge_work(expected.raw.len())?;
        if !reference_projection_matches(projected, expected) {
            return Err(DecodeError::projection());
        }
    }
    Ok((snapshot, budget.report(source.len())))
}

/// Decode one `KN.BuildChunkArchive` payload.
pub fn decode_build_chunk<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<BuildChunkLifecycleSnapshot<'source>, DecodeError> {
    decode_build_chunk_with_report(source, options).map(|(snapshot, _)| snapshot)
}

/// Decode one build chunk and return exact finite usage.
pub fn decode_build_chunk_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(BuildChunkLifecycleSnapshot<'source>, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = parse_build_chunk(source, options, &mut budget)?;
    let view: projection::BuildChunkArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    budget.charge_work(source.len())?;
    let projected_build = view.build.get().map_err(DecodeError::from)?;
    let Some(projected_build) = projected_build.as_ref() else {
        return Err(DecodeError::projection());
    };
    budget.charge_work(snapshot.build.raw.len())?;
    if !reference_projection_matches(projected_build, snapshot.build) {
        return Err(DecodeError::projection());
    }
    let projected_chunk_uuid = view
        .build_chunk_identifier
        .get()
        .map_err(DecodeError::from)?
        .map(|identifier| {
            identifier
                .build_id
                .get()
                .map_err(DecodeError::from)?
                .map(|uuid| {
                    if !uuid.has_lower() || !uuid.has_upper() {
                        Err(DecodeError::projection())
                    } else {
                        Ok(Uuid::new(uuid.lower, uuid.upper))
                    }
                })
                .transpose()
        })
        .transpose()?
        .flatten();
    let projected_build_id = view
        .build_id
        .get()
        .map_err(DecodeError::from)?
        .map(|uuid| {
            if !uuid.has_lower() || !uuid.has_upper() {
                Err(DecodeError::projection())
            } else {
                Ok(Uuid::new(uuid.lower, uuid.upper))
            }
        })
        .transpose()?;
    if let Some(uuid) = snapshot.chunk_identifier {
        budget.charge_work(uuid.raw.len())?;
    }
    if let Some(uuid) = snapshot.build_id {
        budget.charge_work(uuid.raw.len())?;
    }
    if projected_chunk_uuid != snapshot.chunk_identifier.map(|uuid| uuid.uuid())
        || projected_build_id != snapshot.build_id.map(|uuid| uuid.uuid())
    {
        return Err(DecodeError::projection());
    }
    Ok((snapshot, budget.report(source.len())))
}

/// Prepare a source-witnessed slide lifecycle rewrite.
pub fn prepare_slide_lifecycle_rewrite<'source, 'edit>(
    source: &'source [u8],
    edit: SlideLifecycleEdit<'edit>,
    options: DecodeOptions,
) -> Result<PreparedSlideLifecycleRewrite<'source, 'edit>, DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = parse_slide(source, options, &mut budget)?;
    validate_slide_edit(&snapshot, edit, options, &mut budget)?;
    let output_bytes = measure_slide(source, &snapshot, edit, options, &mut budget)?;
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::limited(DecodeLimit::OutputBytes {
            observed: output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    // Reserve the future emission and candidate readback in the same
    // operation ledger before a caller can request the output allocation.
    let estimated_work_bytes = output_bytes
        .checked_mul(4)
        .and_then(|value| value.checked_add(budget.work_bytes))
        .ok_or_else(DecodeError::invalid)?;
    if estimated_work_bytes > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: estimated_work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    Ok(PreparedSlideLifecycleRewrite {
        source,
        edit,
        options,
        snapshot,
        output_bytes,
        estimated_work_bytes,
    })
}

/// Rewrite slide ownership, z-order, build, and build-chunk reference lists.
pub fn rewrite_slide_lifecycle(
    source: &[u8],
    edit: SlideLifecycleEdit<'_>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    rewrite_slide_lifecycle_with_report(source, edit, options).map(|(output, _)| output)
}

/// Rewrite one slide lifecycle payload with a complete source/output report.
pub fn rewrite_slide_lifecycle_with_report(
    source: &[u8],
    edit: SlideLifecycleEdit<'_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    prepare_slide_lifecycle_rewrite(source, edit, options)?.commit()
}

/// Rewrite one build's drawable reference while preserving all other fields.
pub fn rewrite_build(
    source: &[u8],
    edit: BuildLifecycleEdit,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    rewrite_build_with_report(source, edit, options).map(|(output, _)| output)
}

/// Rewrite one build's drawable reference and return exact usage.
pub fn rewrite_build_with_report(
    source: &[u8],
    edit: BuildLifecycleEdit,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = parse_build(source, options, &mut budget)?;
    let current = snapshot
        .drawable
        .ok_or_else(|| DecodeError::missing("KN.BuildArchive.drawable"))?;
    validate_identifier_rewrite(edit.drawable, current.identifier())?;
    if current.identifier() != edit.drawable.source() {
        return Err(DecodeError::unsupported(
            "build drawable rewrite source does not match the witness",
        ));
    }
    let output_bytes = measure_build(source, edit.drawable, options, &mut budget)?;
    ensure_output(output_bytes, options)?;
    charge_rewrite_reserve(&mut budget, output_bytes, options)?;
    let mut output = reserve_output(output_bytes)?;
    emit_build(source, edit.drawable, options, &mut budget, &mut output)?;
    let readback_options =
        options.with_max_message_bytes(options.max_message_bytes.max(output.len()));
    let (readback, report) = decode_build_with_report(&output, readback_options)?;
    if readback.drawable.map(|reference| reference.identifier()) != Some(edit.drawable.target()) {
        return Err(DecodeError::projection());
    }
    let changed = output.as_slice() != source;
    Ok((
        output,
        rewrite_report(source, &budget, &report, output_bytes, changed),
    ))
}

/// Rewrite build and UUID edges in one `KN.BuildChunkArchive` payload.
pub fn rewrite_build_chunk(
    source: &[u8],
    edit: BuildChunkLifecycleEdit,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    rewrite_build_chunk_with_report(source, edit, options).map(|(output, _)| output)
}

/// Rewrite build and UUID edges in one build chunk with exact usage.
pub fn rewrite_build_chunk_with_report(
    source: &[u8],
    edit: BuildChunkLifecycleEdit,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = parse_build_chunk(source, options, &mut budget)?;
    if let Some(rewrite) = edit.build {
        validate_identifier_rewrite(rewrite, snapshot.build.identifier())?;
    }
    if let Some(rewrite) = edit.uuid {
        let source_uuid = snapshot
            .chunk_identifier
            .map(|uuid| uuid.uuid())
            .or_else(|| snapshot.build_id.map(|uuid| uuid.uuid()))
            .ok_or_else(|| DecodeError::missing("KN.BuildChunkArchive.build_id"))?;
        if source_uuid != rewrite.source() {
            return Err(DecodeError::unsupported(
                "build chunk UUID rewrite source does not match the witness",
            ));
        }
        if snapshot
            .chunk_identifier
            .is_some_and(|uuid| uuid.uuid() != source_uuid)
            || snapshot
                .build_id
                .is_some_and(|uuid| uuid.uuid() != source_uuid)
        {
            return Err(DecodeError::unsupported(
                "build chunk contains divergent UUID edges",
            ));
        }
    }
    let output_bytes = measure_build_chunk(source, edit, options, &mut budget)?;
    ensure_output(output_bytes, options)?;
    charge_rewrite_reserve(&mut budget, output_bytes, options)?;
    let mut output = reserve_output(output_bytes)?;
    emit_build_chunk(source, edit, options, &mut budget, &mut output)?;
    let readback_options =
        options.with_max_message_bytes(options.max_message_bytes.max(output.len()));
    let (readback, report) = decode_build_chunk_with_report(&output, readback_options)?;
    if let Some(rewrite) = edit.build {
        if readback.build.identifier() != rewrite.target() {
            return Err(DecodeError::projection());
        }
    }
    if let Some(rewrite) = edit.uuid {
        if readback.chunk_identifier.map(|uuid| uuid.uuid()) != Some(rewrite.target())
            && readback.build_id.map(|uuid| uuid.uuid()) != Some(rewrite.target())
        {
            return Err(DecodeError::projection());
        }
    }
    let changed = output.as_slice() != source;
    Ok((
        output,
        rewrite_report(source, &budget, &report, output_bytes, changed),
    ))
}

#[derive(Debug, Clone, Copy)]
struct Budget {
    options: DecodeOptions,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    scratch_bytes: usize,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Self {
        Self {
            options,
            fields: 0,
            work_bytes: source.len(),
            max_depth: 1,
            allocations: 0,
            scratch_bytes: 0,
        }
    }

    fn report(self, source_bytes: usize) -> DecodeReport {
        DecodeReport {
            input_bytes: source_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            allocations: self.allocations,
            retained_bytes: source_bytes,
            scratch_bytes: self.scratch_bytes,
        }
    }

    fn field(&mut self, bytes: usize, depth: u32) -> Result<(), DecodeError> {
        let fields = self
            .fields
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
        if fields > self.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed: fields,
                maximum: self.options.max_fields,
            }));
        }
        self.fields = fields;
        self.max_depth = self.max_depth.max(depth);
        self.charge_work(bytes.max(1))
    }

    fn charge_work(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self
            .work_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn scratch(&mut self, bytes: usize) -> Result<(), DecodeError> {
        self.scratch_bytes = self
            .scratch_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::invalid)?;
        self.charge_work(bytes)
    }

    fn reserve<T>(
        &mut self,
        vector: &mut Vec<T>,
        additional: usize,
        name: &'static str,
    ) -> Result<(), DecodeError> {
        let before = vector.capacity();
        vector
            .try_reserve(additional)
            .map_err(|_| DecodeError::allocation(name))?;
        let after = vector.capacity();
        if after > before {
            self.allocations = self.allocations.saturating_add(1);
            self.scratch(after.saturating_sub(before).saturating_mul(size_of::<T>()))?;
        }
        Ok(())
    }
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard_limit =
        usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_| DecodeError::invalid())?;
    if options.max_message_bytes > hard_limit || source.len() > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: source.len().max(options.max_message_bytes),
            maximum: hard_limit.min(options.max_message_bytes),
        }));
    }
    if options.max_depth == 0 || options.max_depth > MAX_RECURSION_LIMIT {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: options.max_depth,
            maximum: MAX_RECURSION_LIMIT,
        }));
    }
    if options.max_fields == 0 {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: 1,
            maximum: options.max_fields,
        }));
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct Field<'source> {
    number: u32,
    wire: u8,
    canonical_key: bool,
    canonical_value: bool,
    raw: &'source [u8],
    value: FieldValue<'source>,
}

#[derive(Clone, Copy)]
enum FieldValue<'source> {
    Varint(u64),
    Bytes(&'source [u8]),
    Other,
}

impl<'source> Field<'source> {
    fn varint(self, name: &'static str) -> Result<u64, DecodeError> {
        if self.wire != 0 {
            return Err(DecodeError::wrong_wire(name));
        }
        if !self.canonical_key || !self.canonical_value {
            return Err(DecodeError::noncanonical("field key or varint"));
        }
        match self.value {
            FieldValue::Varint(value) => Ok(value),
            FieldValue::Bytes(_) | FieldValue::Other => Err(DecodeError::projection()),
        }
    }

    fn bytes(self, name: &'static str) -> Result<&'source [u8], DecodeError> {
        if self.wire != 2 {
            return Err(DecodeError::wrong_wire(name));
        }
        if !self.canonical_key || !self.canonical_value {
            return Err(DecodeError::noncanonical("field key or length"));
        }
        match self.value {
            FieldValue::Bytes(value) => Ok(value),
            FieldValue::Varint(_) | FieldValue::Other => Err(DecodeError::projection()),
        }
    }
}

struct Parser<'source, 'budget> {
    remaining: &'source [u8],
    depth: u32,
    options: DecodeOptions,
    budget: &'budget mut Budget,
}

impl<'source, 'budget> Parser<'source, 'budget> {
    fn new(
        source: &'source [u8],
        depth: u32,
        options: DecodeOptions,
        budget: &'budget mut Budget,
    ) -> Result<Self, DecodeError> {
        if depth == 0 || depth > options.max_depth {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: depth,
                maximum: options.max_depth,
            }));
        }
        budget.max_depth = budget.max_depth.max(depth);
        budget.charge_work(source.len())?;
        Ok(Self {
            remaining: source,
            depth,
            options,
            budget,
        })
    }

    fn next(&mut self) -> Result<Option<Field<'source>>, DecodeError> {
        if self.remaining.is_empty() {
            return Ok(None);
        }
        let original = self.remaining;
        let item = parse_item(&mut self.remaining, self.depth, self.options, self.budget)?;
        match item {
            ParseItem::Field(field) => {
                let consumed = original.len().saturating_sub(self.remaining.len());
                Ok(Some(Field {
                    raw: &original[..consumed],
                    ..field
                }))
            },
            ParseItem::EndGroup(_) => {
                Err(DecodeError::unsupported("unexpected protobuf end-group"))
            },
        }
    }
}

enum ParseItem<'source> {
    Field(Field<'source>),
    EndGroup(u32),
}

fn parse_item<'source>(
    input: &mut &'source [u8],
    depth: u32,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ParseItem<'source>, DecodeError> {
    let (tag, canonical_key) = take_varint(input)?;
    let number = u32::try_from(tag >> 3).map_err(|_| DecodeError::invalid())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid());
    }
    let wire = u8::try_from(tag & 7).map_err(|_| DecodeError::invalid())?;
    if wire == 4 {
        budget.field(1, depth)?;
        return Ok(ParseItem::EndGroup(number));
    }
    let canonical_value;
    let value = match wire {
        0 => {
            let (value, canonical) = take_varint(input)?;
            canonical_value = canonical;
            FieldValue::Varint(value)
        },
        1 => {
            let _ = take_exact(input, 8)?;
            canonical_value = true;
            FieldValue::Other
        },
        2 => {
            let (length, canonical) = take_varint(input)?;
            canonical_value = canonical;
            let length = usize::try_from(length).map_err(|_| DecodeError::invalid())?;
            FieldValue::Bytes(take_exact(input, length)?)
        },
        3 => {
            let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
            if child_depth > options.max_depth {
                return Err(DecodeError::limited(DecodeLimit::Nesting {
                    observed: child_depth,
                    maximum: options.max_depth,
                }));
            }
            skip_group(input, number, child_depth, options, budget)?;
            canonical_value = true;
            FieldValue::Other
        },
        5 => {
            let _ = take_exact(input, 4)?;
            canonical_value = true;
            FieldValue::Other
        },
        _ => return Err(DecodeError::invalid()),
    };
    let field = Field {
        number,
        wire,
        canonical_key,
        canonical_value,
        raw: &[],
        value,
    };
    // The caller reconstructs the raw span from its original remaining slice.
    budget.field(1, depth)?;
    Ok(ParseItem::Field(field))
}

fn skip_group(
    input: &mut &[u8],
    expected: u32,
    depth: u32,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    loop {
        let item = parse_item(input, depth, options, budget)?;
        match item {
            ParseItem::Field(_) => {},
            ParseItem::EndGroup(number) if number == expected => return Ok(()),
            ParseItem::EndGroup(_) => return Err(DecodeError::invalid()),
        }
    }
}

fn take_exact<'source>(
    input: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], DecodeError> {
    if input.len() < length {
        return Err(DecodeError::invalid());
    }
    let (selected, remaining) = input.split_at(length);
    *input = remaining;
    Ok(selected)
}

fn take_varint(input: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let mut value = 0u64;
    for index in 0..10 {
        let byte = *input.first().ok_or_else(DecodeError::invalid)?;
        *input = &input[1..];
        if index == 9 && byte > 1 {
            return Err(DecodeError::invalid());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let canonical = encoded_varint_len(value) == index + 1;
            return Ok((value, canonical));
        }
    }
    Err(DecodeError::invalid())
}

fn encoded_varint_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn parse_slide<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<SlideLifecycleSnapshot<'source>, DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    let mut style = None;
    let mut transition = false;
    let mut in_document = None;
    let mut owned_drawables = Vec::new();
    let mut drawables_z_order = Vec::new();
    let mut builds = Vec::new();
    let mut build_chunks = Vec::new();
    while let Some(field) = parser.next()? {
        match field.number {
            SLIDE_STYLE_FIELD => {
                if style.is_some() {
                    return Err(DecodeError::duplicate("KN.SlideArchive.style"));
                }
                let payload = field.bytes("KN.SlideArchive.style")?;
                style = Some(parse_reference(payload, options, parser.budget, 2)?);
            },
            SLIDE_BUILDS_FIELD => {
                let payload = field.bytes("KN.SlideArchive.builds")?;
                let reference = parse_reference(payload, options, parser.budget, 2)?;
                push_unique_reference(
                    &mut builds,
                    reference,
                    options,
                    parser.budget,
                    "slide build references",
                )?;
            },
            SLIDE_DEPRECATED_BUILD_CHUNKS_FIELD => {
                return Err(DecodeError::unsupported(
                    "deprecated inline slide build chunks",
                ));
            },
            SLIDE_TRANSITION_FIELD => {
                if transition {
                    return Err(DecodeError::duplicate("KN.SlideArchive.transition"));
                }
                transition = true;
                let payload = field.bytes("KN.SlideArchive.transition")?;
                parse_transition(payload, options, parser.budget, 2)?;
            },
            SLIDE_OWNED_DRAWABLES_FIELD => {
                let payload = field.bytes("KN.SlideArchive.ownedDrawables")?;
                let reference = parse_reference(payload, options, parser.budget, 2)?;
                push_unique_reference(
                    &mut owned_drawables,
                    reference,
                    options,
                    parser.budget,
                    "slide owned drawable references",
                )?;
            },
            SLIDE_IN_DOCUMENT_FIELD => {
                if in_document.is_some() {
                    return Err(DecodeError::duplicate("KN.SlideArchive.inDocument"));
                }
                in_document = Some(canonical_bool(field.varint("KN.SlideArchive.inDocument")?)?);
            },
            SLIDE_DRAWABLES_Z_ORDER_FIELD => {
                let payload = field.bytes("KN.SlideArchive.drawablesZOrder")?;
                let reference = parse_reference(payload, options, parser.budget, 2)?;
                push_unique_reference(
                    &mut drawables_z_order,
                    reference,
                    options,
                    parser.budget,
                    "slide drawable z-order references",
                )?;
            },
            SLIDE_BUILD_CHUNKS_FIELD => {
                let payload = field.bytes("KN.SlideArchive.buildChunks")?;
                let reference = parse_reference(payload, options, parser.budget, 2)?;
                push_unique_reference(
                    &mut build_chunks,
                    reference,
                    options,
                    parser.budget,
                    "slide build chunk references",
                )?;
            },
            _ => {},
        }
    }
    let style = style.ok_or_else(|| DecodeError::missing("KN.SlideArchive.style"))?;
    if !transition {
        return Err(DecodeError::missing("KN.SlideArchive.transition"));
    }
    let in_document =
        in_document.ok_or_else(|| DecodeError::missing("KN.SlideArchive.inDocument"))?;
    Ok(SlideLifecycleSnapshot {
        source,
        style,
        in_document,
        owned_drawables,
        drawables_z_order,
        builds,
        build_chunks,
    })
}

fn parse_transition(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    let mut parser = Parser::new(source, depth, options, budget)?;
    let mut attributes = false;
    while let Some(field) = parser.next()? {
        if field.number != TRANSITION_ATTRIBUTES_FIELD {
            continue;
        }
        if attributes {
            return Err(DecodeError::duplicate("KN.TransitionArchive.attributes"));
        }
        attributes = true;
        let payload = field.bytes("KN.TransitionArchive.attributes")?;
        consume_message(payload, options, parser.budget, depth + 1)?;
    }
    if attributes {
        Ok(())
    } else {
        Err(DecodeError::missing("KN.TransitionArchive.attributes"))
    }
}

fn parse_build<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<BuildLifecycleSnapshot<'source>, DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    let mut drawable = None;
    let mut delivery = false;
    let mut attributes = false;
    while let Some(field) = parser.next()? {
        match field.number {
            BUILD_DRAWABLE_FIELD => {
                if drawable.is_some() {
                    return Err(DecodeError::duplicate("KN.BuildArchive.drawable"));
                }
                drawable = Some(parse_reference(
                    field.bytes("KN.BuildArchive.drawable")?,
                    options,
                    parser.budget,
                    2,
                )?);
            },
            BUILD_DELIVERY_FIELD => {
                if delivery {
                    return Err(DecodeError::duplicate("KN.BuildArchive.delivery"));
                }
                delivery = true;
                let payload = field.bytes("KN.BuildArchive.delivery")?;
                core::str::from_utf8(payload)
                    .map_err(|_| DecodeError::unsupported("invalid build delivery text"))?;
            },
            BUILD_ATTRIBUTES_FIELD => {
                if attributes {
                    return Err(DecodeError::duplicate("KN.BuildArchive.attributes"));
                }
                attributes = true;
                consume_message(
                    field.bytes("KN.BuildArchive.attributes")?,
                    options,
                    parser.budget,
                    2,
                )?;
            },
            _ => {},
        }
    }
    if !delivery {
        return Err(DecodeError::missing("KN.BuildArchive.delivery"));
    }
    if !attributes {
        return Err(DecodeError::missing("KN.BuildArchive.attributes"));
    }
    Ok(BuildLifecycleSnapshot { source, drawable })
}

fn parse_build_chunk<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<BuildChunkLifecycleSnapshot<'source>, DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    let mut build = None;
    let mut chunk_identifier = None;
    let mut build_id = None;
    let mut automatic = false;
    let mut referent = false;
    while let Some(field) = parser.next()? {
        match field.number {
            CHUNK_BUILD_FIELD => {
                if build.is_some() {
                    return Err(DecodeError::duplicate("KN.BuildChunkArchive.build"));
                }
                build = Some(parse_reference(
                    field.bytes("KN.BuildChunkArchive.build")?,
                    options,
                    parser.budget,
                    2,
                )?);
            },
            CHUNK_AUTOMATIC_FIELD | CHUNK_REFERENT_FIELD => {
                let seen = if field.number == CHUNK_AUTOMATIC_FIELD {
                    &mut automatic
                } else {
                    &mut referent
                };
                if *seen {
                    return Err(DecodeError::duplicate("KN.BuildChunkArchive.boolean edge"));
                }
                *seen = true;
                canonical_bool(field.varint("KN.BuildChunkArchive.bool")?)?;
            },
            CHUNK_IDENTIFIER_FIELD => {
                if chunk_identifier.is_some() {
                    return Err(DecodeError::duplicate(
                        "KN.BuildChunkArchive.buildChunkIdentifier",
                    ));
                }
                let payload = field.bytes("KN.BuildChunkArchive.buildChunkIdentifier")?;
                chunk_identifier =
                    Some(parse_chunk_identifier(payload, options, parser.budget, 2)?);
            },
            CHUNK_BUILD_ID_FIELD => {
                if build_id.is_some() {
                    return Err(DecodeError::duplicate("KN.BuildChunkArchive.buildId"));
                }
                let payload = field.bytes("KN.BuildChunkArchive.buildId")?;
                build_id = Some(parse_uuid(payload, options, parser.budget, 2)?);
            },
            _ => {},
        }
    }
    let build = build.ok_or_else(|| DecodeError::missing("KN.BuildChunkArchive.build"))?;
    if let (Some(chunk), Some(direct)) = (chunk_identifier, build_id) {
        if chunk.uuid() != direct.uuid() {
            return Err(DecodeError::unsupported("build chunk UUID edges disagree"));
        }
    }
    if chunk_identifier.is_none() && build_id.is_none() {
        return Err(DecodeError::missing("KN.BuildChunkArchive.buildId"));
    }
    Ok(BuildChunkLifecycleSnapshot {
        source,
        build,
        chunk_identifier,
        build_id,
    })
}

fn parse_reference<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<Reference<'source>, DecodeError> {
    let mut parser = Parser::new(source, depth, options, budget)?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    while let Some(field) = parser.next()? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::duplicate("TSP.Reference.identifier"));
                }
                let value = field.varint("TSP.Reference.identifier")?;
                if value == 0 {
                    return Err(DecodeError::unsupported("zero reference identifier"));
                }
                identifier = Some(value);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type.is_some() {
                    return Err(DecodeError::duplicate("TSP.Reference.deprecatedType"));
                }
                deprecated_type = Some(canonical_int32(
                    field.varint("TSP.Reference.deprecatedType")?,
                )?);
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if deprecated_is_external.is_some() {
                    return Err(DecodeError::duplicate("TSP.Reference.deprecatedIsExternal"));
                }
                deprecated_is_external = Some(canonical_bool(
                    field.varint("TSP.Reference.deprecatedIsExternal")?,
                )?);
            },
            _ => {},
        }
    }
    if deprecated_is_external == Some(true) {
        return Err(DecodeError::unsupported("external lifecycle reference"));
    }
    Ok(Reference {
        identifier: identifier.ok_or_else(|| DecodeError::missing("TSP.Reference.identifier"))?,
        deprecated_type,
        deprecated_is_external,
        raw: source,
    })
}

fn parse_uuid<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<UuidSnapshot<'source>, DecodeError> {
    let mut parser = Parser::new(source, depth, options, budget)?;
    let mut lower = None;
    let mut upper = None;
    while let Some(field) = parser.next()? {
        match field.number {
            UUID_LOWER_FIELD => {
                if lower.is_some() {
                    return Err(DecodeError::duplicate("TSP.UUID.lower"));
                }
                lower = Some(field.varint("TSP.UUID.lower")?);
            },
            UUID_UPPER_FIELD => {
                if upper.is_some() {
                    return Err(DecodeError::duplicate("TSP.UUID.upper"));
                }
                upper = Some(field.varint("TSP.UUID.upper")?);
            },
            _ => {},
        }
    }
    Ok(UuidSnapshot {
        uuid: Uuid::new(
            lower.ok_or_else(|| DecodeError::missing("TSP.UUID.lower"))?,
            upper.ok_or_else(|| DecodeError::missing("TSP.UUID.upper"))?,
        ),
        raw: source,
    })
}

fn parse_chunk_identifier<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<UuidSnapshot<'source>, DecodeError> {
    let mut parser = Parser::new(source, depth, options, budget)?;
    let mut uuid = None;
    let mut chunk_id = false;
    while let Some(field) = parser.next()? {
        match field.number {
            CHUNK_IDENTIFIER_UUID_FIELD => {
                if uuid.is_some() {
                    return Err(DecodeError::duplicate(
                        "KN.BuildChunkIdentifierArchive.buildId",
                    ));
                }
                uuid = Some(parse_uuid(
                    field.bytes("KN.BuildChunkIdentifierArchive.buildId")?,
                    options,
                    parser.budget,
                    depth + 1,
                )?);
            },
            CHUNK_IDENTIFIER_INDEX_FIELD => {
                if chunk_id {
                    return Err(DecodeError::duplicate(
                        "KN.BuildChunkIdentifierArchive.buildChunkId",
                    ));
                }
                chunk_id = true;
                let _ =
                    canonical_int32(field.varint("KN.BuildChunkIdentifierArchive.buildChunkId")?)?;
            },
            _ => {},
        }
    }
    uuid.ok_or_else(|| DecodeError::missing("KN.BuildChunkIdentifierArchive.buildId"))
}

fn consume_message(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    let mut parser = Parser::new(source, depth, options, budget)?;
    while parser.next()?.is_some() {}
    Ok(())
}

fn push_reference<'source>(
    values: &mut Vec<Reference<'source>>,
    reference: Reference<'source>,
    options: DecodeOptions,
    budget: &mut Budget,
    name: &'static str,
) -> Result<(), DecodeError> {
    if values.len() >= options.max_references {
        return Err(DecodeError::limited(DecodeLimit::References {
            observed: values.len().saturating_add(1),
            maximum: options.max_references,
        }));
    }
    budget.reserve(values, 1, name)?;
    values.push(reference);
    Ok(())
}

fn push_unique_reference<'source>(
    values: &mut Vec<Reference<'source>>,
    reference: Reference<'source>,
    options: DecodeOptions,
    budget: &mut Budget,
    name: &'static str,
) -> Result<(), DecodeError> {
    budget.charge_work(values.len().saturating_mul(size_of::<u64>()))?;
    if values
        .iter()
        .any(|candidate| candidate.identifier() == reference.identifier())
    {
        return Err(DecodeError::unsupported("duplicate slide reference"));
    }
    push_reference(values, reference, options, budget, name)
}

fn canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar")),
    }
}

fn canonical_int32(value: u64) -> Result<i32, DecodeError> {
    if value > i32::MAX as u64 && value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical("int32 scalar"));
    }
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        reason = "The sign-extension range check proves the cast is canonical."
    )]
    Ok(value as i32)
}

fn force_reference_projection(
    expected: Reference<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let view: projection::ReferenceLazyView<'_> = options
        .buffa()
        .decode_lazy_view(expected.raw)
        .map_err(DecodeError::from)?;
    if !reference_projection_matches(&view, expected) {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn reference_projection_matches(
    view: &projection::ReferenceLazyView<'_>,
    expected: Reference<'_>,
) -> bool {
    view.has_identifier()
        && view.identifier == expected.identifier()
        && view.deprecated_type == expected.deprecated_type()
        && view.deprecated_is_external == expected.deprecated_is_external()
}

fn validate_identifier_rewrite(
    rewrite: IdentifierRewrite,
    current: u64,
) -> Result<(), DecodeError> {
    if rewrite.source == 0 || rewrite.target == 0 {
        return Err(DecodeError::unsupported("zero identifier rewrite"));
    }
    if current != rewrite.source {
        return Err(DecodeError::unsupported(
            "identifier rewrite source does not match the witness",
        ));
    }
    Ok(())
}

fn slide_source_reference_count(snapshot: &SlideLifecycleSnapshot<'_>) -> usize {
    snapshot
        .owned_drawables
        .len()
        .saturating_add(snapshot.drawables_z_order.len())
        .saturating_add(snapshot.builds.len())
        .saturating_add(snapshot.build_chunks.len())
}

fn slide_append_reference_count(edit: SlideLifecycleEdit<'_>) -> usize {
    edit.append_owned_drawables
        .len()
        .saturating_add(edit.append_drawables_z_order.len())
        .saturating_add(edit.append_builds.len())
        .saturating_add(edit.append_build_chunks.len())
}

fn charge_slide_lookup_work(
    snapshot: &SlideLifecycleSnapshot<'_>,
    edit: SlideLifecycleEdit<'_>,
    budget: &mut Budget,
    source_passes: usize,
) -> Result<(), DecodeError> {
    let source_count = slide_source_reference_count(snapshot);
    let append_count = slide_append_reference_count(edit);
    let remove_count = edit.remove_identifiers.len();
    let remap_count = edit.remap_identifiers.len();
    let source_lookup_count = remove_count
        .saturating_add(remap_count)
        .saturating_add(append_count);
    let append_lookup_count = source_count
        .saturating_add(remove_count)
        .saturating_add(remap_count);
    let source_work = source_count
        .saturating_mul(source_lookup_count)
        .saturating_mul(source_passes);
    let append_work = append_count.saturating_mul(append_lookup_count);
    budget.charge_work(source_work.saturating_add(append_work))
}

fn validate_slide_edit(
    snapshot: &SlideLifecycleSnapshot<'_>,
    edit: SlideLifecycleEdit<'_>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    if edit.remove_identifiers.len() > options.max_references {
        return Err(DecodeError::limited(DecodeLimit::References {
            observed: edit.remove_identifiers.len(),
            maximum: options.max_references,
        }));
    }
    if edit.remap_identifiers.len() > options.max_references {
        return Err(DecodeError::limited(DecodeLimit::References {
            observed: edit.remap_identifiers.len(),
            maximum: options.max_references,
        }));
    }
    charge_slide_lookup_work(snapshot, edit, budget, 1)?;
    reject_duplicate_ids(edit.remove_identifiers, budget)?;
    for &identifier in edit.remove_identifiers {
        if identifier == 0 {
            return Err(DecodeError::unsupported("zero removed identifier"));
        }
    }
    budget.charge_work(
        edit.remap_identifiers
            .len()
            .saturating_mul(edit.remap_identifiers.len())
            .saturating_mul(size_of::<IdentifierRewrite>()),
    )?;
    for (index, rewrite) in edit.remap_identifiers.iter().copied().enumerate() {
        validate_identifier_rewrite(rewrite, rewrite.source())?;
        for prior in &edit.remap_identifiers[..index] {
            if prior.source() == rewrite.source()
                || prior.target() == rewrite.target()
                || prior.source() == rewrite.target()
                || prior.target() == rewrite.source()
            {
                return Err(DecodeError::unsupported("ambiguous identifier remap"));
            }
        }
    }
    let source_lists = [
        snapshot.owned_drawables.as_slice(),
        snapshot.drawables_z_order.as_slice(),
        snapshot.builds.as_slice(),
        snapshot.build_chunks.as_slice(),
    ];
    for &identifier in edit.remove_identifiers {
        if !source_lists.iter().any(|list| {
            list.iter()
                .any(|reference| reference.identifier() == identifier)
        }) {
            return Err(DecodeError::unsupported(
                "removed identifier is absent from every slide lifecycle list",
            ));
        }
    }
    let lists = [
        (
            snapshot.owned_drawables.as_slice(),
            edit.append_owned_drawables,
        ),
        (
            snapshot.drawables_z_order.as_slice(),
            edit.append_drawables_z_order,
        ),
        (snapshot.builds.as_slice(), edit.append_builds),
        (snapshot.build_chunks.as_slice(), edit.append_build_chunks),
    ];
    for (existing, appended) in lists {
        if appended.len() > options.max_references
            || existing.len().saturating_add(appended.len()) > options.max_references
        {
            return Err(DecodeError::limited(DecodeLimit::References {
                observed: existing.len().saturating_add(appended.len()),
                maximum: options.max_references,
            }));
        }
        reject_duplicate_ids(appended, budget)?;
        for &identifier in appended {
            if identifier == 0 {
                return Err(DecodeError::unsupported("zero appended identifier"));
            }
            if existing
                .iter()
                .any(|reference| reference.identifier() == identifier)
            {
                return Err(DecodeError::unsupported(
                    "appended identifier duplicates an existing reference",
                ));
            }
            if edit.remove_identifiers.contains(&identifier) {
                return Err(DecodeError::unsupported(
                    "appended identifier is also removed",
                ));
            }
            if edit
                .remap_identifiers
                .iter()
                .any(|rewrite| rewrite.target() == identifier)
            {
                return Err(DecodeError::unsupported(
                    "appended identifier conflicts with a remap target",
                ));
            }
        }
        // The remap source/target checks are list-local.  A package owner
        // separately proves global object-identity uniqueness.
        budget.charge_work(
            existing
                .len()
                .saturating_mul(appended.len())
                .saturating_mul(8),
        )?;
    }
    Ok(())
}

fn reject_duplicate_ids(values: &[u64], budget: &mut Budget) -> Result<(), DecodeError> {
    for (index, &value) in values.iter().enumerate() {
        budget.charge_work(index.saturating_mul(size_of::<u64>()))?;
        if values[..index].contains(&value) {
            return Err(DecodeError::unsupported("duplicate edit identifier"));
        }
    }
    Ok(())
}

fn find_remap(remaps: &[IdentifierRewrite], identifier: u64) -> Option<u64> {
    remaps
        .iter()
        .find(|rewrite| rewrite.source() == identifier)
        .map(|rewrite| rewrite.target())
}

fn should_remove(removals: &[u64], identifier: u64) -> bool {
    removals.contains(&identifier)
}

fn measure_slide(
    source: &[u8],
    snapshot: &SlideLifecycleSnapshot<'_>,
    edit: SlideLifecycleEdit<'_>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    charge_slide_lookup_work(snapshot, edit, budget, 1)?;
    let mut parser = Parser::new(source, 1, options, budget)?;
    let mut output_bytes = 0usize;
    while let Some(field) = parser.next()? {
        let list = match field.number {
            SLIDE_OWNED_DRAWABLES_FIELD => Some(edit.append_owned_drawables),
            SLIDE_DRAWABLES_Z_ORDER_FIELD => Some(edit.append_drawables_z_order),
            SLIDE_BUILDS_FIELD => Some(edit.append_builds),
            SLIDE_BUILD_CHUNKS_FIELD => Some(edit.append_build_chunks),
            _ => None,
        };
        let replacement = if list.is_some() {
            let payload = field.bytes("slide repeated reference")?;
            let reference = parse_reference(payload, options, parser.budget, 2)?;
            if should_remove(edit.remove_identifiers, reference.identifier()) {
                0
            } else if let Some(target) = find_remap(edit.remap_identifiers, reference.identifier())
            {
                length_field_size(
                    field.number,
                    measure_reference(payload, Some(target), options, parser.budget)?,
                )?
            } else {
                field.raw.len()
            }
        } else {
            field.raw.len()
        };
        output_bytes = checked_add_output(output_bytes, replacement, options)?;
    }
    for (field, values) in [
        (SLIDE_OWNED_DRAWABLES_FIELD, edit.append_owned_drawables),
        (SLIDE_DRAWABLES_Z_ORDER_FIELD, edit.append_drawables_z_order),
        (SLIDE_BUILDS_FIELD, edit.append_builds),
        (SLIDE_BUILD_CHUNKS_FIELD, edit.append_build_chunks),
    ] {
        for &identifier in values {
            output_bytes = checked_add_output(
                output_bytes,
                length_field_size(field, reference_encoded_len(identifier))?,
                options,
            )?;
            budget.charge_work(size_of::<u64>())?;
        }
    }
    Ok(output_bytes)
}

fn emit_slide(
    source: &[u8],
    snapshot: &SlideLifecycleSnapshot<'_>,
    edit: SlideLifecycleEdit<'_>,
    options: DecodeOptions,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    charge_slide_lookup_work(snapshot, edit, budget, 1)?;
    let mut parser = Parser::new(source, 1, options, budget)?;
    while let Some(field) = parser.next()? {
        let selected = matches!(
            field.number,
            SLIDE_OWNED_DRAWABLES_FIELD
                | SLIDE_DRAWABLES_Z_ORDER_FIELD
                | SLIDE_BUILDS_FIELD
                | SLIDE_BUILD_CHUNKS_FIELD
        );
        if !selected {
            output.extend_from_slice(field.raw);
            continue;
        }
        let payload = field.bytes("slide repeated reference")?;
        let reference = parse_reference(payload, options, parser.budget, 2)?;
        if should_remove(edit.remove_identifiers, reference.identifier()) {
            continue;
        }
        if let Some(target) = find_remap(edit.remap_identifiers, reference.identifier()) {
            emit_rewritten_length_field(
                field.number,
                payload,
                Some(target),
                options,
                parser.budget,
                output,
            )?;
        } else {
            output.extend_from_slice(field.raw);
        }
    }
    for (field, values) in [
        (SLIDE_OWNED_DRAWABLES_FIELD, edit.append_owned_drawables),
        (SLIDE_DRAWABLES_Z_ORDER_FIELD, edit.append_drawables_z_order),
        (SLIDE_BUILDS_FIELD, edit.append_builds),
        (SLIDE_BUILD_CHUNKS_FIELD, edit.append_build_chunks),
    ] {
        for &identifier in values {
            emit_reference_field(field, identifier, output);
        }
    }
    Ok(())
}

fn same_slide_semantics(
    candidate: &SlideLifecycleSnapshot<'_>,
    source: &SlideLifecycleSnapshot<'_>,
    edit: SlideLifecycleEdit<'_>,
    budget: &mut Budget,
) -> Result<bool, DecodeError> {
    charge_slide_lookup_work(source, edit, budget, 1)?;
    if candidate.style.identifier() != source.style.identifier()
        || candidate.in_document != source.in_document
    {
        return Ok(false);
    }
    let expected = |values: &[Reference<'_>], appended: &[u64]| {
        let mut ids = Vec::new();
        ids.try_reserve(values.len().saturating_add(appended.len()))
            .map_err(|_| DecodeError::allocation("slide semantic readback"))?;
        for value in values {
            if should_remove(edit.remove_identifiers, value.identifier()) {
                continue;
            }
            ids.push(
                find_remap(edit.remap_identifiers, value.identifier())
                    .unwrap_or(value.identifier()),
            );
        }
        ids.extend_from_slice(appended);
        Ok::<Vec<u64>, DecodeError>(ids)
    };
    let check = |actual: &[Reference<'_>], original: &[Reference<'_>], appended: &[u64]| {
        let expected_ids = expected(original, appended)?;
        Ok::<bool, DecodeError>(
            actual
                .iter()
                .map(|reference| reference.identifier())
                .eq(expected_ids),
        )
    };
    Ok(check(
        &candidate.owned_drawables,
        &source.owned_drawables,
        edit.append_owned_drawables,
    )? && check(
        &candidate.drawables_z_order,
        &source.drawables_z_order,
        edit.append_drawables_z_order,
    )? && check(&candidate.builds, &source.builds, edit.append_builds)?
        && check(
            &candidate.build_chunks,
            &source.build_chunks,
            edit.append_build_chunks,
        )?)
}

fn checked_add_output(
    current: usize,
    addition: usize,
    options: DecodeOptions,
) -> Result<usize, DecodeError> {
    let output = current
        .checked_add(addition)
        .ok_or_else(DecodeError::invalid)?;
    if output > options.max_output_bytes {
        return Err(DecodeError::limited(DecodeLimit::OutputBytes {
            observed: output,
            maximum: options.max_output_bytes,
        }));
    }
    Ok(output)
}

fn length_field_size(field: u32, payload: usize) -> Result<usize, DecodeError> {
    encoded_varint_len((u64::from(field) << 3) | 2)
        .checked_add(encoded_varint_len(
            u64::try_from(payload).map_err(|_| DecodeError::invalid())?,
        ))
        .and_then(|value| value.checked_add(payload))
        .ok_or_else(DecodeError::invalid)
}

fn reference_encoded_len(identifier: u64) -> usize {
    varint_field_size(REFERENCE_IDENTIFIER_FIELD, identifier)
}

fn varint_field_size(field: u32, value: u64) -> usize {
    encoded_varint_len(u64::from(field) << 3).saturating_add(encoded_varint_len(value))
}

fn reserve_output(size: usize) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(size)
        .map_err(|_| DecodeError::allocation("lifecycle output"))?;
    Ok(output)
}

fn ensure_output(output_bytes: usize, options: DecodeOptions) -> Result<(), DecodeError> {
    if output_bytes > options.max_output_bytes {
        Err(DecodeError::limited(DecodeLimit::OutputBytes {
            observed: output_bytes,
            maximum: options.max_output_bytes,
        }))
    } else {
        Ok(())
    }
}

fn charge_rewrite_reserve(
    budget: &mut Budget,
    output_bytes: usize,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let reserve = output_bytes
        .checked_mul(4)
        .ok_or_else(DecodeError::invalid)?;
    budget.charge_work(reserve)?;
    if budget.work_bytes > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: budget.work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    Ok(())
}

fn measure_reference(
    source: &[u8],
    replacement: Option<u64>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    let mut output = 0usize;
    while let Some(field) = parser.next()? {
        let size = if field.number == REFERENCE_IDENTIFIER_FIELD {
            let value = field.varint("TSP.Reference.identifier")?;
            let replacement = replacement.unwrap_or(value);
            varint_field_size(REFERENCE_IDENTIFIER_FIELD, replacement)
        } else {
            field.raw.len()
        };
        output = output.checked_add(size).ok_or_else(DecodeError::invalid)?;
    }
    Ok(output)
}

fn emit_rewritten_reference(
    source: &[u8],
    replacement: Option<u64>,
    options: DecodeOptions,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    while let Some(field) = parser.next()? {
        if field.number == REFERENCE_IDENTIFIER_FIELD {
            let value = field.varint("TSP.Reference.identifier")?;
            emit_varint_field(
                REFERENCE_IDENTIFIER_FIELD,
                replacement.unwrap_or(value),
                output,
            );
        } else {
            output.extend_from_slice(field.raw);
        }
    }
    Ok(())
}

fn emit_rewritten_length_field(
    field: u32,
    source: &[u8],
    replacement: Option<u64>,
    options: DecodeOptions,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let payload_size = measure_reference(source, replacement, options, budget)?;
    emit_key_and_length(field, payload_size, output);
    emit_rewritten_reference(source, replacement, options, budget, output)
}

fn emit_reference_field(field: u32, identifier: u64, output: &mut Vec<u8>) {
    let payload_size = reference_encoded_len(identifier);
    emit_key_and_length(field, payload_size, output);
    emit_varint_field(REFERENCE_IDENTIFIER_FIELD, identifier, output);
}

fn measure_build(
    source: &[u8],
    rewrite: IdentifierRewrite,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    let mut output = 0usize;
    while let Some(field) = parser.next()? {
        let size = if field.number == BUILD_DRAWABLE_FIELD {
            let payload = field.bytes("KN.BuildArchive.drawable")?;
            length_field_size(
                BUILD_DRAWABLE_FIELD,
                measure_reference(payload, Some(rewrite.target()), options, parser.budget)?,
            )?
        } else {
            field.raw.len()
        };
        output = output.checked_add(size).ok_or_else(DecodeError::invalid)?;
    }
    Ok(output)
}

fn emit_build(
    source: &[u8],
    rewrite: IdentifierRewrite,
    options: DecodeOptions,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    while let Some(field) = parser.next()? {
        if field.number != BUILD_DRAWABLE_FIELD {
            output.extend_from_slice(field.raw);
            continue;
        }
        let payload = field.bytes("KN.BuildArchive.drawable")?;
        let payload_size =
            measure_reference(payload, Some(rewrite.target()), options, parser.budget)?;
        emit_key_and_length(BUILD_DRAWABLE_FIELD, payload_size, output);
        emit_rewritten_reference(
            payload,
            Some(rewrite.target()),
            options,
            parser.budget,
            output,
        )?;
    }
    Ok(())
}

fn measure_build_chunk(
    source: &[u8],
    edit: BuildChunkLifecycleEdit,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    let mut output = 0usize;
    while let Some(field) = parser.next()? {
        let size = match field.number {
            CHUNK_BUILD_FIELD if edit.build.is_some() => {
                let payload = field.bytes("KN.BuildChunkArchive.build")?;
                length_field_size(
                    CHUNK_BUILD_FIELD,
                    measure_reference(
                        payload,
                        Some(edit.build.ok_or_else(DecodeError::invalid)?.target()),
                        options,
                        parser.budget,
                    )?,
                )?
            },
            CHUNK_IDENTIFIER_FIELD if edit.uuid.is_some() => {
                let payload = field.bytes("KN.BuildChunkArchive.buildChunkIdentifier")?;
                length_field_size(
                    CHUNK_IDENTIFIER_FIELD,
                    measure_chunk_identifier(payload, edit.uuid, options, parser.budget)?,
                )?
            },
            CHUNK_BUILD_ID_FIELD if edit.uuid.is_some() => {
                let payload = field.bytes("KN.BuildChunkArchive.buildId")?;
                length_field_size(
                    CHUNK_BUILD_ID_FIELD,
                    measure_uuid(payload, edit.uuid, options, parser.budget)?,
                )?
            },
            _ => field.raw.len(),
        };
        output = output.checked_add(size).ok_or_else(DecodeError::invalid)?;
    }
    Ok(output)
}

fn emit_build_chunk(
    source: &[u8],
    edit: BuildChunkLifecycleEdit,
    options: DecodeOptions,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    while let Some(field) = parser.next()? {
        match field.number {
            CHUNK_BUILD_FIELD if edit.build.is_some() => {
                let payload = field.bytes("KN.BuildChunkArchive.build")?;
                let rewrite = edit.build.ok_or_else(DecodeError::invalid)?;
                let payload_size =
                    measure_reference(payload, Some(rewrite.target()), options, parser.budget)?;
                emit_key_and_length(CHUNK_BUILD_FIELD, payload_size, output);
                emit_rewritten_reference(
                    payload,
                    Some(rewrite.target()),
                    options,
                    parser.budget,
                    output,
                )?;
            },
            CHUNK_IDENTIFIER_FIELD if edit.uuid.is_some() => {
                let payload = field.bytes("KN.BuildChunkArchive.buildChunkIdentifier")?;
                let payload_size =
                    measure_chunk_identifier(payload, edit.uuid, options, parser.budget)?;
                emit_key_and_length(CHUNK_IDENTIFIER_FIELD, payload_size, output);
                emit_chunk_identifier(payload, edit.uuid, options, parser.budget, output)?;
            },
            CHUNK_BUILD_ID_FIELD if edit.uuid.is_some() => {
                let payload = field.bytes("KN.BuildChunkArchive.buildId")?;
                let payload_size = measure_uuid(payload, edit.uuid, options, parser.budget)?;
                emit_key_and_length(CHUNK_BUILD_ID_FIELD, payload_size, output);
                emit_uuid(payload, edit.uuid, options, parser.budget, output)?;
            },
            _ => output.extend_from_slice(field.raw),
        }
    }
    Ok(())
}

fn measure_uuid(
    source: &[u8],
    rewrite: Option<UuidRewrite>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    let mut output = 0usize;
    while let Some(field) = parser.next()? {
        let size = match field.number {
            UUID_LOWER_FIELD => {
                let value = field.varint("TSP.UUID.lower")?;
                let value = rewrite.map_or(value, |rewrite| rewrite.target().lower());
                varint_field_size(UUID_LOWER_FIELD, value)
            },
            UUID_UPPER_FIELD => {
                let value = field.varint("TSP.UUID.upper")?;
                let value = rewrite.map_or(value, |rewrite| rewrite.target().upper());
                varint_field_size(UUID_UPPER_FIELD, value)
            },
            _ => field.raw.len(),
        };
        output = output.checked_add(size).ok_or_else(DecodeError::invalid)?;
    }
    Ok(output)
}

fn emit_uuid(
    source: &[u8],
    rewrite: Option<UuidRewrite>,
    options: DecodeOptions,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    while let Some(field) = parser.next()? {
        match field.number {
            UUID_LOWER_FIELD => {
                let value = field.varint("TSP.UUID.lower")?;
                emit_varint_field(
                    UUID_LOWER_FIELD,
                    rewrite.map_or(value, |rewrite| rewrite.target().lower()),
                    output,
                );
            },
            UUID_UPPER_FIELD => {
                let value = field.varint("TSP.UUID.upper")?;
                emit_varint_field(
                    UUID_UPPER_FIELD,
                    rewrite.map_or(value, |rewrite| rewrite.target().upper()),
                    output,
                );
            },
            _ => output.extend_from_slice(field.raw),
        }
    }
    Ok(())
}

fn measure_chunk_identifier(
    source: &[u8],
    rewrite: Option<UuidRewrite>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    let mut output = 0usize;
    while let Some(field) = parser.next()? {
        let size = if field.number == CHUNK_IDENTIFIER_UUID_FIELD {
            let payload = field.bytes("KN.BuildChunkIdentifierArchive.buildId")?;
            length_field_size(
                CHUNK_IDENTIFIER_UUID_FIELD,
                measure_uuid(payload, rewrite, options, parser.budget)?,
            )?
        } else {
            field.raw.len()
        };
        output = output.checked_add(size).ok_or_else(DecodeError::invalid)?;
    }
    Ok(output)
}

fn emit_chunk_identifier(
    source: &[u8],
    rewrite: Option<UuidRewrite>,
    options: DecodeOptions,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    while let Some(field) = parser.next()? {
        if field.number != CHUNK_IDENTIFIER_UUID_FIELD {
            output.extend_from_slice(field.raw);
            continue;
        }
        let payload = field.bytes("KN.BuildChunkIdentifierArchive.buildId")?;
        let payload_size = measure_uuid(payload, rewrite, options, parser.budget)?;
        emit_key_and_length(CHUNK_IDENTIFIER_UUID_FIELD, payload_size, output);
        emit_uuid(payload, rewrite, options, parser.budget, output)?;
    }
    Ok(())
}

fn emit_key_and_length(field: u32, payload: usize, output: &mut Vec<u8>) {
    emit_varint((u64::from(field) << 3) | 2, output);
    emit_varint(payload as u64, output);
}

fn emit_varint_field(field: u32, value: u64, output: &mut Vec<u8>) {
    emit_varint(u64::from(field) << 3, output);
    emit_varint(value, output);
}

fn emit_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn rewrite_report(
    source: &[u8],
    budget: &Budget,
    readback: &DecodeReport,
    output_bytes: usize,
    changed: bool,
) -> RewriteReport {
    RewriteReport {
        input_bytes: source.len(),
        output_bytes,
        fields: budget.fields.saturating_add(readback.fields),
        work_bytes: budget.work_bytes.saturating_add(readback.work_bytes),
        max_depth: budget.max_depth.max(readback.max_depth),
        allocations: 1,
        retained_bytes: source.len().saturating_add(output_bytes),
        scratch_bytes: budget.scratch_bytes.saturating_add(readback.scratch_bytes),
        changed,
    }
}

#[cfg(test)]
#[path = "keynote_media_lifecycle_codec_tests.rs"]
mod tests;
