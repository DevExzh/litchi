//! Bounded, source-preserving ownership for the outer `KN.ThemeArchive`.
//!
//! The `TSS.ThemeArchive` value in field one is deliberately opaque. This
//! module owns the Keynote envelope, projects the selected references and
//! scalars through a private Buffa lazy view, and keeps every unselected field
//! as an exact source span. Classic-theme, timing-curve, and
//! motion-background fields are framed but not decoded here.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict source pass intentionally precedes the private Buffa view."
)]

use std::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_theme_generated::LitchiIwaKeynoteThemeProjection as projection;

const THEME_SUPER_FIELD: u32 = 1;
const THEME_TEMPLATES_FIELD: u32 = 2;
const THEME_UUID_FIELD: u32 = 3;
const THEME_DEFAULT_TEMPLATE_FIELD: u32 = 5;
const THEME_DEFAULT_TEMPLATE_REFERENCE_FIELD: u32 = 6;
const THEME_DEFAULT_TEMPLATE_GUESS_FIELD: u32 = 7;
const THEME_LIVE_VIDEO_FIELD: u32 = 9;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MAX_RECURSION_LIMIT: u32 = 64;
const MAX_AUTOMATIC_LIMIT: usize = 512 * 1024 * 1024;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Which finite resource ceiling rejected a source or transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecodeLimit {
    /// Complete input bytes inspected by the transaction.
    InputBytes,
    /// Encoded fields visited by the strict and replay passes.
    Fields,
    /// Aggregate source and projection work charged before execution.
    WorkBytes,
    /// Candidate output bytes.
    OutputBytes,
    /// Logical output allocations.
    Allocations,
    /// Source and candidate bytes retained together.
    RetainedBytes,
    /// Candidate staging bytes.
    ScratchBytes,
    /// Maximum nested protobuf depth.
    Nesting,
    /// Number of repeated template records.
    TemplateReferences,
}

impl fmt::Display for DecodeLimit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = match self {
            Self::InputBytes => "input bytes",
            Self::Fields => "fields",
            Self::WorkBytes => "work bytes",
            Self::OutputBytes => "output bytes",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::Nesting => "nesting",
            Self::TemplateReferences => "template references",
        };
        formatter.write_str(name)
    }
}

/// Stable category for a bounded outer-theme failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DecodeErrorKind {
    /// Malformed protobuf framing or a selected-field wire violation.
    InvalidFormat(String),
    /// A finite resource ceiling was exceeded.
    Resource {
        /// Rejected resource category.
        limit: DecodeLimit,
        /// Observed or requested amount.
        observed: usize,
        /// Applied ceiling.
        maximum: usize,
    },
    /// A required selected field was absent.
    MissingRequired(String),
    /// A singular selected field appeared more than once.
    DuplicateSingular(String),
    /// A selected varint, key, or scalar used a non-canonical form.
    NonCanonical(String),
    /// The private Buffa projection disagreed with the strict source pass.
    Projection,
    /// A destination buffer could not be reserved.
    Allocation { amount: usize },
}

/// Error returned by a bounded outer-theme operation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

impl DecodeError {
    fn invalid(message: impl Into<String>) -> Self {
        Self {
            kind: DecodeErrorKind::InvalidFormat(message.into()),
        }
    }

    fn resource(limit: DecodeLimit, observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Resource {
                limit,
                observed,
                maximum,
            },
        }
    }

    fn missing(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::MissingRequired(field.into()),
        }
    }

    fn duplicate(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateSingular(field.into()),
        }
    }

    fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason.into()),
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    const fn allocation(amount: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation { amount },
        }
    }

    /// Return the stable failure category.
    #[must_use]
    pub const fn kind(&self) -> &DecodeErrorKind {
        &self.kind
    }

    /// Return the rejected finite resource, if any.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Resource { limit, .. } => Some(limit),
            _ => None,
        }
    }

    /// Return the requested allocation amount, if reservation failed.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        match self.kind {
            DecodeErrorKind::Allocation { amount } => Some(amount),
            _ => None,
        }
    }

    /// Return the missing selected field, if any.
    #[must_use]
    pub fn missing_required_field(&self) -> Option<&str> {
        match &self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field.as_str()),
            _ => None,
        }
    }

    /// Return the duplicate selected field, if any.
    #[must_use]
    pub fn duplicate_singular_field(&self) -> Option<&str> {
        match &self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field.as_str()),
            _ => None,
        }
    }

    /// Return the selected-field canonicality reason, if any.
    #[must_use]
    pub fn noncanonical_reason(&self) -> Option<&str> {
        match &self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason.as_str()),
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::InvalidFormat(message) => formatter.write_str(message),
            DecodeErrorKind::Resource {
                limit,
                observed,
                maximum,
            } => write!(
                formatter,
                "Keynote theme {limit} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "Keynote theme required field {field} is absent")
            },
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(
                    formatter,
                    "Keynote theme field {field} occurs more than once"
                )
            },
            DecodeErrorKind::NonCanonical(reason) => formatter.write_str(reason),
            DecodeErrorKind::Projection => {
                formatter.write_str("Keynote theme Buffa projection disagreed with strict source")
            },
            DecodeErrorKind::Allocation { amount } => {
                write!(
                    formatter,
                    "Keynote theme allocation of {amount} bytes failed"
                )
            },
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self::invalid(format!("Buffa Keynote theme projection: {error}"))
    }
}

/// Finite limits for one theme source or rewrite transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_output_bytes: usize,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
    max_template_references: usize,
}

impl DecodeOptions {
    /// Build an explicit finite policy.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_output_bytes: max_message_bytes,
            max_allocations: 1,
            max_retained_bytes: max_message_bytes,
            max_scratch_bytes: max_message_bytes,
            max_template_references: max_fields,
        }
    }

    /// Build a meaningful finite profile from a known source width.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(8).clamp(1, MAX_AUTOMATIC_LIMIT),
            bytes.saturating_mul(32).clamp(1, MAX_AUTOMATIC_LIMIT),
            MAX_RECURSION_LIMIT,
        )
        .with_max_output_bytes(bytes.saturating_mul(2).clamp(1, MAX_AUTOMATIC_LIMIT))
        .with_max_retained_bytes(bytes.saturating_mul(4).clamp(1, MAX_AUTOMATIC_LIMIT))
        .with_max_scratch_bytes(bytes.saturating_mul(2).clamp(1, MAX_AUTOMATIC_LIMIT))
        .with_max_template_references(bytes.saturating_mul(2).clamp(1, MAX_AUTOMATIC_LIMIT))
    }

    /// Replace the input-byte ceiling.
    #[must_use]
    pub const fn with_max_input_bytes(self, maximum: usize) -> Self {
        self.with_max_message_bytes(maximum)
    }

    /// Replace the input-byte ceiling.
    #[must_use]
    pub const fn with_max_message_bytes(mut self, maximum: usize) -> Self {
        self.max_message_bytes = maximum;
        self
    }

    /// Replace the encoded-field ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the aggregate work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Replace the candidate-output ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the logical output-allocation ceiling.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }

    /// Replace the source-plus-candidate retained-byte ceiling.
    #[must_use]
    pub const fn with_max_retained_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_bytes = maximum;
        self
    }

    /// Replace the candidate scratch-byte ceiling.
    #[must_use]
    pub const fn with_max_scratch_bytes(mut self, maximum: usize) -> Self {
        self.max_scratch_bytes = maximum;
        self
    }

    /// Replace the repeated template-reference ceiling.
    #[must_use]
    pub const fn with_max_template_references(mut self, maximum: usize) -> Self {
        self.max_template_references = maximum;
        self
    }

    /// Replace the protobuf nesting ceiling.
    #[must_use]
    pub const fn with_max_depth(mut self, maximum: u32) -> Self {
        self.recursion_limit = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }

    fn internal_for(source_len: usize) -> Self {
        Self::new(
            source_len.max(1),
            MAX_AUTOMATIC_LIMIT,
            MAX_AUTOMATIC_LIMIT,
            MAX_RECURSION_LIMIT,
        )
        .with_max_output_bytes(MAX_AUTOMATIC_LIMIT)
        .with_max_retained_bytes(MAX_AUTOMATIC_LIMIT)
        .with_max_scratch_bytes(MAX_AUTOMATIC_LIMIT)
        .with_max_template_references(MAX_AUTOMATIC_LIMIT)
    }
}

impl Default for DecodeOptions {
    fn default() -> Self {
        Self::new(
            64 * 1024 * 1024,
            1_000_000,
            256 * 1024 * 1024,
            MAX_RECURSION_LIMIT,
        )
        .with_max_output_bytes(128 * 1024 * 1024)
        .with_max_retained_bytes(192 * 1024 * 1024)
        .with_max_scratch_bytes(128 * 1024 * 1024)
        .with_max_template_references(1_000_000)
    }
}

/// Exact charged consumption of a successful borrowed decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    template_references: usize,
}

impl DecodeReport {
    /// Input bytes inspected.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Alias for input bytes.
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.input_bytes
    }

    /// Strict and replay field visits charged before execution.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate source and projection work charged before execution.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum nested depth observed.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Logical allocations performed by the borrowed decode.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source bytes retained by the snapshot.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Scratch bytes allocated by the borrowed decode.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Number of template records validated.
    #[must_use]
    pub const fn template_references(self) -> usize {
        self.template_references
    }
}

/// Exact finite requirements computed before a fresh outer theme allocation.
#[derive(Debug)]
pub struct ThemeArchiveEncodeRequirements<'source> {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    template_references: usize,
    candidate_fields: usize,
    candidate_work: usize,
    source: ThemeArchiveBuilder<'source>,
}

impl ThemeArchiveEncodeRequirements<'_> {
    /// Exact encoded outer theme length.
    #[must_use]
    pub const fn output_bytes(&self) -> usize {
        self.output_bytes
    }

    /// Aggregate measured field visits, including candidate validation.
    #[must_use]
    pub const fn fields(&self) -> usize {
        self.fields
    }

    /// Aggregate measured work, including candidate validation.
    #[must_use]
    pub const fn work_bytes(&self) -> usize {
        self.work_bytes
    }

    /// Maximum nested reference depth.
    #[must_use]
    pub const fn max_depth(&self) -> u32 {
        self.max_depth
    }

    /// Logical output allocations.
    #[must_use]
    pub const fn allocations(&self) -> usize {
        self.allocations
    }

    /// Source and candidate bytes retained during the encode/readback phase.
    #[must_use]
    pub const fn retained_bytes(&self) -> usize {
        self.retained_bytes
    }

    /// Candidate staging bytes.
    #[must_use]
    pub const fn scratch_bytes(&self) -> usize {
        self.scratch_bytes
    }

    /// Number of template references admitted by the candidate decoder.
    #[must_use]
    pub const fn template_references(&self) -> usize {
        self.template_references
    }

    /// Convert this measurement into exact ceilings for [`ThemeArchiveBuilder::encode`].
    #[must_use]
    pub const fn exact_options(&self) -> DecodeOptions {
        DecodeOptions::new(
            if self.output_bytes == 0 {
                1
            } else {
                self.output_bytes
            },
            self.fields,
            self.work_bytes,
            if self.max_depth == 0 {
                1
            } else {
                self.max_depth
            },
        )
        .with_max_output_bytes(self.output_bytes)
        .with_max_allocations(if self.allocations == 0 {
            1
        } else {
            self.allocations
        })
        .with_max_retained_bytes(self.retained_bytes)
        .with_max_scratch_bytes(self.scratch_bytes)
        .with_max_template_references(self.template_references)
    }
}

/// One selected TSP reference projected from the Keynote envelope.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ThemeReferenceSnapshot {
    /// Required native object identifier.
    pub identifier: u64,
    /// Optional deprecated type marker.
    pub deprecated_type: Option<i32>,
    /// Optional deprecated external marker.
    pub deprecated_is_external: Option<bool>,
}

impl ThemeReferenceSnapshot {
    /// Construct a reference containing only its required identifier.
    #[must_use]
    pub const fn new(identifier: u64) -> Self {
        Self {
            identifier,
            deprecated_type: None,
            deprecated_is_external: None,
        }
    }

    /// Encode this reference as canonical nested protobuf bytes.
    ///
    /// # Errors
    ///
    /// A typed invalid-source [`DecodeError`] when the identifier is zero
    /// (record 0368a: no panicking twin).
    pub fn encode(self) -> Result<Vec<u8>, DecodeError> {
        if self.identifier == 0 {
            return Err(DecodeError::invalid(
                "TSP.Reference.identifier must be non-zero",
            ));
        }
        let mut output = Vec::new();
        append_reference_payload(&mut output, self);
        Ok(output)
    }
}

/// Borrowed, source-preserving view of one complete outer theme archive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ThemeArchiveSnapshot<'source> {
    source: &'source [u8],
    super_bytes: &'source [u8],
    uuid: Option<&'source str>,
    template_references: usize,
    default_template_slide_node: Option<ThemeReferenceSnapshot>,
    default_template_slide_node_reference: Option<ThemeReferenceSnapshot>,
    default_template_slide_node_is_our_best_guess: Option<bool>,
    live_video_source_collection: Option<ThemeReferenceSnapshot>,
}

impl<'source> ThemeArchiveSnapshot<'source> {
    /// Return the complete source bytes.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }

    /// Return the opaque TSS theme payload from field one.
    #[must_use]
    pub const fn super_bytes(self) -> &'source [u8] {
        self.super_bytes
    }

    /// Return the optional UUID.
    #[must_use]
    pub const fn uuid(self) -> Option<&'source str> {
        self.uuid
    }

    /// Return template references in native source order.
    #[must_use]
    pub const fn templates(self) -> ThemeReferenceFields<'source> {
        ThemeReferenceFields {
            source: self.source,
            cursor: 0,
            remaining: self.template_references,
        }
    }

    /// Return the number of template records.
    #[must_use]
    pub const fn template_count(self) -> usize {
        self.template_references
    }

    /// Return the optional default slide-node reference.
    #[must_use]
    pub const fn default_template_slide_node(self) -> Option<ThemeReferenceSnapshot> {
        self.default_template_slide_node
    }

    /// Return the optional default slide-node reference object.
    #[must_use]
    pub const fn default_template_slide_node_reference(self) -> Option<ThemeReferenceSnapshot> {
        self.default_template_slide_node_reference
    }

    /// Return the optional default-reference heuristic flag.
    #[must_use]
    pub const fn default_template_slide_node_is_our_best_guess(self) -> Option<bool> {
        self.default_template_slide_node_is_our_best_guess
    }

    /// Return the optional live-video source collection reference.
    #[must_use]
    pub const fn live_video_source_collection(self) -> Option<ThemeReferenceSnapshot> {
        self.live_video_source_collection
    }
}

/// Bounded source iterator for field-two template references.
#[derive(Debug, Clone, Copy)]
pub struct ThemeReferenceFields<'source> {
    source: &'source [u8],
    cursor: usize,
    remaining: usize,
}

impl Iterator for ThemeReferenceFields<'_> {
    type Item = ThemeReferenceSnapshot;

    fn next(&mut self) -> Option<Self::Item> {
        while self.remaining != 0 {
            let field = next_field_unchecked(self.source, &mut self.cursor)?;
            if field.number != THEME_TEMPLATES_FIELD {
                continue;
            }
            let payload = field.payload?;
            let reference = reference_snapshot_unchecked(payload)?;
            self.remaining -= 1;
            return Some(reference);
        }
        None
    }
}

/// A source-preserving mutation of selected outer fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ThemeArchiveWrite<'source> {
    super_bytes: Patch<&'source [u8]>,
    templates: Option<&'source [ThemeReferenceSnapshot]>,
    uuid: Patch<&'source str>,
    default_template_slide_node: Patch<ThemeReferenceSnapshot>,
    default_template_slide_node_reference: Patch<ThemeReferenceSnapshot>,
    default_template_slide_node_is_our_best_guess: Patch<bool>,
    live_video_source_collection: Patch<ThemeReferenceSnapshot>,
}

impl<'source> ThemeArchiveWrite<'source> {
    /// Build a no-op mutation preserving every source span.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            super_bytes: Patch::Preserve,
            templates: None,
            uuid: Patch::Preserve,
            default_template_slide_node: Patch::Preserve,
            default_template_slide_node_reference: Patch::Preserve,
            default_template_slide_node_is_our_best_guess: Patch::Preserve,
            live_video_source_collection: Patch::Preserve,
        }
    }

    /// Replace the opaque TSS theme payload.
    #[must_use]
    pub const fn with_super(mut self, payload: &'source [u8]) -> Self {
        self.super_bytes = Patch::Replace(payload);
        self
    }

    /// Replace all template references. An empty slice removes field two.
    #[must_use]
    pub const fn with_templates(mut self, templates: &'source [ThemeReferenceSnapshot]) -> Self {
        self.templates = Some(templates);
        self
    }

    /// Replace the UUID string.
    #[must_use]
    pub const fn with_uuid(mut self, uuid: &'source str) -> Self {
        self.uuid = Patch::Replace(uuid);
        self
    }

    /// Remove the UUID field.
    #[must_use]
    pub const fn clear_uuid(mut self) -> Self {
        self.uuid = Patch::Clear;
        self
    }

    /// Replace the default slide-node reference.
    #[must_use]
    pub const fn with_default_template_slide_node(
        mut self,
        reference: ThemeReferenceSnapshot,
    ) -> Self {
        self.default_template_slide_node = Patch::Replace(reference);
        self
    }

    /// Remove the default slide-node reference.
    #[must_use]
    pub const fn clear_default_template_slide_node(mut self) -> Self {
        self.default_template_slide_node = Patch::Clear;
        self
    }

    /// Replace the default slide-node reference object.
    #[must_use]
    pub const fn with_default_template_slide_node_reference(
        mut self,
        reference: ThemeReferenceSnapshot,
    ) -> Self {
        self.default_template_slide_node_reference = Patch::Replace(reference);
        self
    }

    /// Remove the default slide-node reference object.
    #[must_use]
    pub const fn clear_default_template_slide_node_reference(mut self) -> Self {
        self.default_template_slide_node_reference = Patch::Clear;
        self
    }

    /// Replace the default-reference heuristic flag.
    #[must_use]
    pub const fn with_default_template_slide_node_is_our_best_guess(mut self, value: bool) -> Self {
        self.default_template_slide_node_is_our_best_guess = Patch::Replace(value);
        self
    }

    /// Remove the default-reference heuristic flag.
    #[must_use]
    pub const fn clear_default_template_slide_node_is_our_best_guess(mut self) -> Self {
        self.default_template_slide_node_is_our_best_guess = Patch::Clear;
        self
    }

    /// Replace the live-video source collection reference.
    #[must_use]
    pub const fn with_live_video_source_collection(
        mut self,
        reference: ThemeReferenceSnapshot,
    ) -> Self {
        self.live_video_source_collection = Patch::Replace(reference);
        self
    }

    /// Remove the live-video source collection reference.
    #[must_use]
    pub const fn clear_live_video_source_collection(mut self) -> Self {
        self.live_video_source_collection = Patch::Clear;
        self
    }
}

impl Default for ThemeArchiveWrite<'_> {
    fn default() -> Self {
        Self::new()
    }
}

/// Bounded constructor for a fresh outer theme archive.
#[derive(Debug, Clone, Copy)]
pub struct ThemeArchiveBuilder<'source> {
    super_bytes: Option<&'source [u8]>,
    templates: &'source [ThemeReferenceSnapshot],
    uuid: Option<&'source str>,
    default_template_slide_node: Option<ThemeReferenceSnapshot>,
    default_template_slide_node_reference: Option<ThemeReferenceSnapshot>,
    default_template_slide_node_is_our_best_guess: Option<bool>,
    live_video_source_collection: Option<ThemeReferenceSnapshot>,
}

impl<'source> ThemeArchiveBuilder<'source> {
    /// Build an empty constructor; `super` must be supplied before encoding.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            super_bytes: None,
            templates: &[],
            uuid: None,
            default_template_slide_node: None,
            default_template_slide_node_reference: None,
            default_template_slide_node_is_our_best_guess: None,
            live_video_source_collection: None,
        }
    }

    /// Set the opaque required TSS theme payload.
    #[must_use]
    pub const fn with_super(mut self, payload: &'source [u8]) -> Self {
        self.super_bytes = Some(payload);
        self
    }

    /// Set all template references in source order.
    #[must_use]
    pub const fn with_templates(mut self, templates: &'source [ThemeReferenceSnapshot]) -> Self {
        self.templates = templates;
        self
    }

    /// Set the optional UUID.
    #[must_use]
    pub const fn with_uuid(mut self, uuid: &'source str) -> Self {
        self.uuid = Some(uuid);
        self
    }

    /// Set the optional default slide-node reference.
    #[must_use]
    pub const fn with_default_template_slide_node(
        mut self,
        reference: ThemeReferenceSnapshot,
    ) -> Self {
        self.default_template_slide_node = Some(reference);
        self
    }

    /// Set the optional default slide-node reference object.
    #[must_use]
    pub const fn with_default_template_slide_node_reference(
        mut self,
        reference: ThemeReferenceSnapshot,
    ) -> Self {
        self.default_template_slide_node_reference = Some(reference);
        self
    }

    /// Set the optional default-reference heuristic flag.
    #[must_use]
    pub const fn with_default_template_slide_node_is_our_best_guess(mut self, value: bool) -> Self {
        self.default_template_slide_node_is_our_best_guess = Some(value);
        self
    }

    /// Set the optional live-video source collection reference.
    #[must_use]
    pub const fn with_live_video_source_collection(
        mut self,
        reference: ThemeReferenceSnapshot,
    ) -> Self {
        self.live_video_source_collection = Some(reference);
        self
    }

    /// Measure a fresh canonical outer archive before allocating its output.
    pub fn encode_requirements(
        self,
        options: DecodeOptions,
    ) -> Result<ThemeArchiveEncodeRequirements<'source>, DecodeError> {
        validate_options(options)?;
        let super_bytes = self
            .super_bytes
            .ok_or_else(|| DecodeError::missing("KN.ThemeArchive.super"))?;
        validate_reference_slice(self.templates, options)?;

        let mut output_len = encoded_field_len(THEME_SUPER_FIELD, super_bytes.len())?;
        let mut output_fields = 1usize;
        let mut template_payload_bytes = 0usize;
        let mut selected_nested_bytes = 0usize;
        let mut has_nested_reference = false;
        for reference in self.templates {
            output_len = checked_add(
                output_len,
                encoded_reference_field_len(THEME_TEMPLATES_FIELD, *reference),
                options.max_output_bytes,
            )?;
            output_fields = output_fields
                .checked_add(1 + reference_field_count(*reference))
                .ok_or_else(|| {
                    DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
                })?;
            template_payload_bytes = template_payload_bytes
                .checked_add(encoded_reference_payload_len(*reference))
                .ok_or_else(|| {
                    DecodeError::resource(
                        DecodeLimit::WorkBytes,
                        usize::MAX,
                        options.max_work_bytes,
                    )
                })?;
            has_nested_reference = true;
        }
        if let Some(uuid) = self.uuid {
            output_len = checked_add(
                output_len,
                encoded_field_len(THEME_UUID_FIELD, uuid.len())?,
                options.max_output_bytes,
            )?;
            output_fields = output_fields.checked_add(1).ok_or_else(|| {
                DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
            })?;
        }
        for (field, reference) in [
            (
                THEME_DEFAULT_TEMPLATE_FIELD,
                self.default_template_slide_node,
            ),
            (
                THEME_DEFAULT_TEMPLATE_REFERENCE_FIELD,
                self.default_template_slide_node_reference,
            ),
            (THEME_LIVE_VIDEO_FIELD, self.live_video_source_collection),
        ] {
            if let Some(reference) = reference {
                validate_reference_identifier(reference)?;
                output_len = checked_add(
                    output_len,
                    encoded_reference_field_len(field, reference),
                    options.max_output_bytes,
                )?;
                output_fields = output_fields
                    .checked_add(1 + reference_field_count(reference))
                    .ok_or_else(|| {
                        DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
                    })?;
                selected_nested_bytes = selected_nested_bytes
                    .checked_add(encoded_reference_payload_len(reference))
                    .ok_or_else(|| {
                        DecodeError::resource(
                            DecodeLimit::WorkBytes,
                            usize::MAX,
                            options.max_work_bytes,
                        )
                    })?;
                has_nested_reference = true;
            }
        }
        if self.default_template_slide_node_is_our_best_guess.is_some() {
            output_len = checked_add(output_len, 2, options.max_output_bytes)?;
            output_fields = output_fields.checked_add(1).ok_or_else(|| {
                DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
            })?;
        }

        let candidate_fields = output_fields.checked_mul(3).ok_or_else(|| {
            DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
        })?;
        let candidate_work = output_len
            .checked_mul(3)
            .and_then(|value| {
                template_payload_bytes
                    .checked_mul(2)
                    .and_then(|amount| value.checked_add(amount))
            })
            .and_then(|value| {
                selected_nested_bytes
                    .checked_mul(2)
                    .and_then(|amount| value.checked_add(amount))
            })
            .ok_or_else(|| {
                DecodeError::resource(DecodeLimit::WorkBytes, usize::MAX, options.max_work_bytes)
            })?;
        let work = output_len
            .checked_add(output_fields)
            .and_then(|value| value.checked_add(candidate_work))
            .ok_or_else(|| {
                DecodeError::resource(DecodeLimit::WorkBytes, usize::MAX, options.max_work_bytes)
            })?;
        let fields = output_fields.checked_add(candidate_fields).ok_or_else(|| {
            DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
        })?;
        let max_depth = if has_nested_reference { 2 } else { 1 };
        if options.recursion_limit < max_depth {
            return Err(DecodeError::resource(
                DecodeLimit::Nesting,
                max_depth as usize,
                options.recursion_limit as usize,
            ));
        }
        admit_output(
            &options, output_len, fields, work, 1, output_len, output_len,
        )?;
        Ok(ThemeArchiveEncodeRequirements {
            output_bytes: output_len,
            fields,
            work_bytes: work,
            max_depth,
            allocations: 1,
            retained_bytes: output_len,
            scratch_bytes: output_len,
            template_references: self.templates.len(),
            candidate_fields,
            candidate_work,
            source: self,
        })
    }

    /// Encode a fresh canonical outer archive after all resource admission.
    pub fn encode(self, options: DecodeOptions) -> Result<Vec<u8>, DecodeError> {
        let requirements = self.encode_requirements(options)?;
        requirements.encode()
    }
}

impl<'source> ThemeArchiveEncodeRequirements<'source> {
    /// Consume this prepared token and encode its exact borrowed builder.
    ///
    /// The requirement token is intentionally not `Clone` or `Copy`. Its
    /// lifetime keeps every source slice borrowed until this consuming call,
    /// and execution cannot be redirected to a builder whose content was not
    /// measured.
    pub fn encode(self) -> Result<Vec<u8>, DecodeError> {
        let ThemeArchiveEncodeRequirements {
            output_bytes,
            max_depth,
            candidate_fields,
            candidate_work,
            source,
            ..
        } = self;
        let output_len = output_bytes;
        let super_bytes = source
            .super_bytes
            .ok_or_else(|| DecodeError::missing("KN.ThemeArchive.super"))?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|_| DecodeError::allocation(output_len))?;
        append_length_delimited(&mut output, THEME_SUPER_FIELD, super_bytes);
        for reference in source.templates {
            append_reference_field(&mut output, THEME_TEMPLATES_FIELD, *reference);
        }
        if let Some(uuid) = source.uuid {
            append_length_delimited(&mut output, THEME_UUID_FIELD, uuid.as_bytes());
        }
        if let Some(reference) = source.default_template_slide_node {
            append_reference_field(&mut output, THEME_DEFAULT_TEMPLATE_FIELD, reference);
        }
        if let Some(reference) = source.default_template_slide_node_reference {
            append_reference_field(
                &mut output,
                THEME_DEFAULT_TEMPLATE_REFERENCE_FIELD,
                reference,
            );
        }
        if let Some(value) = source.default_template_slide_node_is_our_best_guess {
            append_varint_field(
                &mut output,
                THEME_DEFAULT_TEMPLATE_GUESS_FIELD,
                u64::from(value),
            );
        }
        if let Some(reference) = source.live_video_source_collection {
            append_reference_field(&mut output, THEME_LIVE_VIDEO_FIELD, reference);
        }
        debug_assert_eq!(output.len(), output_len);
        let candidate_options = DecodeOptions::new(
            output.len().max(1),
            candidate_fields,
            candidate_work,
            max_depth,
        )
        .with_max_output_bytes(output.len())
        .with_max_allocations(0)
        .with_max_retained_bytes(output.len())
        .with_max_scratch_bytes(0)
        .with_max_template_references(source.templates.len());
        decode_theme_archive(&output, candidate_options)?;
        Ok(output)
    }
}

impl Default for ThemeArchiveBuilder<'_> {
    fn default() -> Self {
        Self::new()
    }
}

/// Exact aggregate costs charged by a source-preserving rewrite.
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
    /// Input source bytes.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Candidate output bytes.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Aggregate field visits charged before allocation.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate work bytes charged before allocation.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum nested depth required by the transaction.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Logical output allocations.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source and output bytes retained together.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Candidate staging bytes.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Whether the candidate differs byte-for-byte from the source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// Exact finite requirements computed before a rewrite allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    /// Candidate output bytes.
    pub output_bytes: usize,
    /// Aggregate field visits across source, emission, and readback charges.
    pub fields: usize,
    /// Aggregate work bytes charged before execution.
    pub work_bytes: usize,
    /// Maximum nested protobuf depth.
    pub max_depth: u32,
    /// Logical output allocations.
    pub allocations: usize,
    /// Source and candidate bytes retained together.
    pub retained_bytes: usize,
    /// Candidate staging bytes.
    pub scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
    /// Convert requirements to exact execution limits.
    #[must_use]
    pub const fn exact(self) -> RewriteExecutionLimits {
        RewriteExecutionLimits {
            output_bytes: self.output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
        }
    }
}

/// Caller ceilings checked before the candidate `Vec` is allocated.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    /// Candidate output bytes.
    pub output_bytes: usize,
    /// Aggregate field visits.
    pub fields: usize,
    /// Aggregate work bytes.
    pub work_bytes: usize,
    /// Maximum nested protobuf depth.
    pub max_depth: u32,
    /// Logical output allocations.
    pub allocations: usize,
    /// Source and candidate bytes retained together.
    pub retained_bytes: usize,
    /// Candidate staging bytes.
    pub scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    /// Convert exact requirements to exact limits.
    #[must_use]
    pub const fn exact(requirements: RewriteExecutionRequirements) -> Self {
        requirements.exact()
    }

    /// Replace the output-byte ceiling.
    #[must_use]
    pub const fn with_output_bytes(mut self, value: usize) -> Self {
        self.output_bytes = value;
        self
    }

    /// Replace the field ceiling.
    #[must_use]
    pub const fn with_fields(mut self, value: usize) -> Self {
        self.fields = value;
        self
    }

    /// Replace the work-byte ceiling.
    #[must_use]
    pub const fn with_work_bytes(mut self, value: usize) -> Self {
        self.work_bytes = value;
        self
    }

    /// Replace the maximum-depth ceiling.
    #[must_use]
    pub const fn with_max_depth(mut self, value: u32) -> Self {
        self.max_depth = value;
        self
    }

    /// Replace the allocation ceiling.
    #[must_use]
    pub const fn with_allocations(mut self, value: usize) -> Self {
        self.allocations = value;
        self
    }

    /// Replace the retained-byte ceiling.
    #[must_use]
    pub const fn with_retained_bytes(mut self, value: usize) -> Self {
        self.retained_bytes = value;
        self
    }

    /// Replace the scratch-byte ceiling.
    #[must_use]
    pub const fn with_scratch_bytes(mut self, value: usize) -> Self {
        self.scratch_bytes = value;
        self
    }
}

/// Candidate bytes plus exact rewrite accounting.
#[derive(Debug, PartialEq, Eq)]
pub struct RewriteOutput {
    bytes: Vec<u8>,
    report: RewriteReport,
}

impl RewriteOutput {
    /// Borrow candidate bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Consume the output and return candidate bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    /// Return exact rewrite accounting.
    #[must_use]
    pub const fn report(&self) -> RewriteReport {
        self.report
    }
}

/// Prepared source-preserving rewrite. No candidate allocation occurs until
/// [`PreparedThemeArchiveRewrite::execute`] admits the exact limits.
#[derive(Debug)]
pub struct PreparedThemeArchiveRewrite<'source, 'write> {
    source: &'source [u8],
    write: ThemeArchiveWrite<'write>,
    before: ThemeArchiveSnapshot<'source>,
    plan: RewritePlan,
    requirements: RewriteExecutionRequirements,
}

impl<'source, 'write> PreparedThemeArchiveRewrite<'source, 'write> {
    /// Return exact requirements computed before output allocation.
    #[must_use]
    pub const fn execution_requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute the prepared rewrite after checking every caller ceiling.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        ensure_requirements(self.requirements, limits)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| DecodeError::allocation(self.requirements.output_bytes))?;
        let emit_options = DecodeOptions::new(
            self.source.len().max(1),
            self.requirements.fields,
            self.requirements.work_bytes,
            self.requirements.max_depth.max(1),
        );
        emit_rewrite(self.source, self.write, &mut output, emit_options)?;
        if output.len() != self.plan.output_len {
            return Err(DecodeError::invalid(
                "Keynote theme rewrite size precharge disagreed with emission",
            ));
        }
        let candidate_options = DecodeOptions::new(
            output.len().max(1),
            self.requirements.fields,
            self.requirements.work_bytes,
            self.requirements.max_depth.max(1),
        )
        .with_max_output_bytes(self.requirements.output_bytes)
        .with_max_allocations(0)
        .with_max_retained_bytes(self.requirements.retained_bytes)
        .with_max_scratch_bytes(0)
        .with_max_template_references(self.requirements.fields);
        let after = decode_theme_archive(&output, candidate_options)?;
        if !same_selected_fields(self.before, after, self.write) {
            return Err(DecodeError::projection());
        }
        let report = RewriteReport {
            input_bytes: self.source.len(),
            output_bytes: output.len(),
            fields: self.requirements.fields,
            work_bytes: self.requirements.work_bytes,
            max_depth: self.requirements.max_depth,
            allocations: self.requirements.allocations,
            retained_bytes: self.requirements.retained_bytes,
            scratch_bytes: self.requirements.scratch_bytes,
            changed: output.as_slice() != self.source,
        };
        Ok(RewriteOutput {
            bytes: output,
            report,
        })
    }
}

/// Decode one outer Keynote theme with a borrowed snapshot.
pub fn decode_theme_archive(
    source: &[u8],
    options: DecodeOptions,
) -> Result<ThemeArchiveSnapshot<'_>, DecodeError> {
    decode_theme_archive_with_report(source, options).map(|(snapshot, _)| snapshot)
}

/// Decode one outer Keynote theme and return its exact precharged report.
pub fn decode_theme_archive_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(ThemeArchiveSnapshot<'_>, DecodeReport), DecodeError> {
    validate_options(options)?;
    ensure_input(&options, source.len())?;
    let mut budget = Budget::new(options);
    let preflight = preflight_theme(source, options, &mut budget)?;
    let strict_fields = budget.fields;

    // Account for one post-decode template iterator replay before invoking
    // the first lazy view. The iterator remains allocation-free after return.
    budget.reserve_fields(strict_fields)?;
    budget.charge_work(
        source
            .len()
            .checked_add(preflight.template_payload_bytes)
            .ok_or_else(|| {
                DecodeError::resource(DecodeLimit::WorkBytes, usize::MAX, options.max_work_bytes)
            })?,
    )?;

    // Buffa sees only the bounded projection. Its selected lazy fields are
    // forced below, while all unselected source spans stay caller-owned.
    budget.reserve_fields(strict_fields)?;
    budget.charge_work(
        source
            .len()
            .checked_add(preflight.selected_nested_bytes)
            .ok_or_else(|| {
                DecodeError::resource(DecodeLimit::WorkBytes, usize::MAX, options.max_work_bytes)
            })?,
    )?;
    force_projection(source, &preflight, options)?;

    let snapshot = ThemeArchiveSnapshot {
        source,
        super_bytes: preflight.super_bytes,
        uuid: preflight.uuid,
        template_references: preflight.template_references,
        default_template_slide_node: preflight.default_template_slide_node,
        default_template_slide_node_reference: preflight.default_template_slide_node_reference,
        default_template_slide_node_is_our_best_guess: preflight
            .default_template_slide_node_is_our_best_guess,
        live_video_source_collection: preflight.live_video_source_collection,
    };
    Ok((
        snapshot,
        budget.report(source.len(), preflight.template_references),
    ))
}

/// Rewrite selected outer fields while copying all unselected source spans.
pub fn rewrite_theme_archive(
    source: &[u8],
    write: ThemeArchiveWrite<'_>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let prepared = prepare_theme_archive_rewrite(source, write, options)?;
    let requirements = prepared.execution_requirements();
    Ok(prepared
        .execute(RewriteExecutionLimits::exact(requirements))?
        .into_bytes())
}

/// Prepare a source-preserving rewrite and compute all execution requirements
/// before a candidate buffer is allocated.
pub fn prepare_theme_archive_rewrite<'source, 'write>(
    source: &'source [u8],
    write: ThemeArchiveWrite<'write>,
    options: DecodeOptions,
) -> Result<PreparedThemeArchiveRewrite<'source, 'write>, DecodeError> {
    let (before, decode_report) = decode_theme_archive_with_report(source, options)?;
    validate_write(&write, options)?;
    let plan = measure_rewrite(source, write, options)?;

    let candidate_fields = plan.output_fields.checked_mul(3).ok_or_else(|| {
        DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
    })?;
    let comparison_fields = plan
        .source_fields
        .checked_add(plan.output_fields)
        .ok_or_else(|| {
            DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
        })?;
    let fields = decode_report
        .fields()
        .checked_add(plan.source_fields)
        .and_then(|value| value.checked_add(plan.source_fields))
        .and_then(|value| value.checked_add(plan.output_fields))
        .and_then(|value| value.checked_add(candidate_fields))
        .and_then(|value| value.checked_add(comparison_fields))
        .ok_or_else(|| {
            DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
        })?;
    let candidate_work = plan
        .output_len
        .checked_mul(3)
        .and_then(|value| {
            plan.output_template_payload_bytes
                .checked_mul(2)
                .and_then(|amount| value.checked_add(amount))
        })
        .and_then(|value| {
            plan.output_selected_nested_bytes
                .checked_mul(2)
                .and_then(|amount| value.checked_add(amount))
        })
        .ok_or_else(|| {
            DecodeError::resource(DecodeLimit::WorkBytes, usize::MAX, options.max_work_bytes)
        })?;
    let comparison_work = source
        .len()
        .checked_add(plan.source_template_payload_bytes)
        .and_then(|value| {
            plan.output_len
                .checked_add(plan.output_template_payload_bytes)
                .and_then(|amount| value.checked_add(amount))
        })
        .ok_or_else(|| {
            DecodeError::resource(DecodeLimit::WorkBytes, usize::MAX, options.max_work_bytes)
        })?;
    let work = decode_report
        .work_bytes()
        .checked_add(source.len())
        .and_then(|value| value.checked_add(source.len()))
        .and_then(|value| value.checked_add(plan.output_len))
        .and_then(|value| value.checked_add(plan.output_fields))
        .and_then(|value| value.checked_add(candidate_work))
        .and_then(|value| value.checked_add(comparison_work))
        .ok_or_else(|| {
            DecodeError::resource(DecodeLimit::WorkBytes, usize::MAX, options.max_work_bytes)
        })?;
    let retained = source.len().checked_add(plan.output_len).ok_or_else(|| {
        DecodeError::resource(
            DecodeLimit::RetainedBytes,
            usize::MAX,
            options.max_retained_bytes,
        )
    })?;
    let candidate_depth =
        if plan.output_template_payload_bytes != 0 || plan.output_selected_nested_bytes != 0 {
            2
        } else {
            1
        };
    let max_depth = decode_report.max_depth().max(candidate_depth);
    if options.recursion_limit < max_depth {
        return Err(DecodeError::resource(
            DecodeLimit::Nesting,
            max_depth as usize,
            options.recursion_limit as usize,
        ));
    }
    admit_output(
        &options,
        plan.output_len,
        fields,
        work,
        1,
        retained,
        plan.output_len,
    )?;

    Ok(PreparedThemeArchiveRewrite {
        source,
        write,
        before,
        plan,
        requirements: RewriteExecutionRequirements {
            output_bytes: plan.output_len,
            fields,
            work_bytes: work,
            max_depth,
            allocations: 1,
            retained_bytes: retained,
            scratch_bytes: plan.output_len,
        },
    })
}

/// Rewrite selected outer fields and return the precharged transaction report.
pub fn rewrite_theme_archive_with_report(
    source: &[u8],
    write: ThemeArchiveWrite<'_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let prepared = prepare_theme_archive_rewrite(source, write, options)?;
    let requirements = prepared.execution_requirements();
    let output = prepared.execute(RewriteExecutionLimits::exact(requirements))?;
    let report = output.report();
    Ok((output.into_bytes(), report))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Patch<T> {
    Preserve,
    Clear,
    Replace(T),
}

#[derive(Debug, Clone, Copy)]
struct Preflight<'source> {
    super_bytes: &'source [u8],
    uuid: Option<&'source str>,
    template_references: usize,
    template_payload_bytes: usize,
    selected_nested_bytes: usize,
    default_template_slide_node: Option<ThemeReferenceSnapshot>,
    default_template_slide_node_reference: Option<ThemeReferenceSnapshot>,
    default_template_slide_node_is_our_best_guess: Option<bool>,
    live_video_source_collection: Option<ThemeReferenceSnapshot>,
}

#[derive(Debug, Clone, Copy)]
struct WireField<'source> {
    number: u32,
    wire_type: u8,
    payload: Option<&'source [u8]>,
    varint: Option<u64>,
    raw: &'source [u8],
    tag_canonical: bool,
    value_canonical: bool,
}

#[derive(Debug, Clone, Copy)]
struct Budget {
    options: DecodeOptions,
    fields: usize,
    work: usize,
    max_depth: u32,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            options,
            fields: 0,
            work: 0,
            max_depth: 0,
        }
    }

    fn field(&mut self, depth: u32) -> Result<(), DecodeError> {
        self.fields = self.fields.checked_add(1).ok_or_else(|| {
            DecodeError::resource(DecodeLimit::Fields, usize::MAX, self.options.max_fields)
        })?;
        if self.fields > self.options.max_fields {
            return Err(DecodeError::resource(
                DecodeLimit::Fields,
                self.fields,
                self.options.max_fields,
            ));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    fn reserve_fields(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self.fields.checked_add(amount).ok_or_else(|| {
            DecodeError::resource(DecodeLimit::Fields, usize::MAX, self.options.max_fields)
        })?;
        if observed > self.options.max_fields {
            return Err(DecodeError::resource(
                DecodeLimit::Fields,
                observed,
                self.options.max_fields,
            ));
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self.work.checked_add(amount).ok_or_else(|| {
            DecodeError::resource(
                DecodeLimit::WorkBytes,
                usize::MAX,
                self.options.max_work_bytes,
            )
        })?;
        if observed > self.options.max_work_bytes {
            return Err(DecodeError::resource(
                DecodeLimit::WorkBytes,
                observed,
                self.options.max_work_bytes,
            ));
        }
        self.work = observed;
        Ok(())
    }

    fn report(self, input_bytes: usize, template_references: usize) -> DecodeReport {
        DecodeReport {
            input_bytes,
            fields: self.fields,
            work_bytes: self.work,
            max_depth: self.max_depth,
            allocations: 0,
            retained_bytes: input_bytes,
            scratch_bytes: 0,
            template_references,
        }
    }
}

fn preflight_theme<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Preflight<'source>, DecodeError> {
    budget.charge_work(source.len())?;
    let mut cursor = 0usize;
    let mut super_bytes = None;
    let mut uuid = None;
    let mut templates = 0usize;
    let mut template_payload_bytes = 0usize;
    let mut selected_nested_bytes = 0usize;
    let mut default_template_slide_node = None;
    let mut default_template_slide_node_reference = None;
    let mut default_template_slide_node_is_our_best_guess = None;
    let mut live_video_source_collection = None;
    let mut seen = 0u16;
    while let Some(field) = next_field(source, &mut cursor, 1, budget)? {
        match field.number {
            THEME_SUPER_FIELD => {
                mark_singular(&mut seen, 0, "KN.ThemeArchive.super")?;
                require_length_delimited(&field, "KN.ThemeArchive.super")?;
                require_canonical(&field, "KN.ThemeArchive.super")?;
                super_bytes = field.payload;
            },
            THEME_TEMPLATES_FIELD => {
                require_length_delimited(&field, "KN.ThemeArchive.templates")?;
                require_canonical(&field, "KN.ThemeArchive.templates")?;
                let payload = field.payload.unwrap_or_default();
                templates = templates.checked_add(1).ok_or_else(|| {
                    DecodeError::resource(
                        DecodeLimit::TemplateReferences,
                        usize::MAX,
                        options.max_template_references,
                    )
                })?;
                if templates > options.max_template_references {
                    return Err(DecodeError::resource(
                        DecodeLimit::TemplateReferences,
                        templates,
                        options.max_template_references,
                    ));
                }
                template_payload_bytes = template_payload_bytes
                    .checked_add(payload.len())
                    .ok_or_else(|| {
                        DecodeError::resource(
                            DecodeLimit::WorkBytes,
                            usize::MAX,
                            options.max_work_bytes,
                        )
                    })?;
                preflight_reference(payload, 2, budget)?;
            },
            THEME_UUID_FIELD => {
                mark_singular(&mut seen, 2, "KN.ThemeArchive.uuid")?;
                require_length_delimited(&field, "KN.ThemeArchive.uuid")?;
                require_canonical(&field, "KN.ThemeArchive.uuid")?;
                uuid = Some(
                    str::from_utf8(field.payload.unwrap_or_default())
                        .map_err(|_| DecodeError::invalid("KN.ThemeArchive.uuid is not UTF-8"))?,
                );
            },
            THEME_DEFAULT_TEMPLATE_FIELD => {
                mark_singular(&mut seen, 4, "KN.ThemeArchive.default_template_slide_node")?;
                let payload = selected_reference_payload(
                    &field,
                    "KN.ThemeArchive.default_template_slide_node",
                )?;
                selected_nested_bytes = selected_nested_bytes
                    .checked_add(payload.len())
                    .ok_or_else(DecodeError::projection)?;
                default_template_slide_node = Some(preflight_reference(payload, 2, budget)?);
            },
            THEME_DEFAULT_TEMPLATE_REFERENCE_FIELD => {
                mark_singular(
                    &mut seen,
                    5,
                    "KN.ThemeArchive.default_template_slide_node_reference",
                )?;
                let payload = selected_reference_payload(
                    &field,
                    "KN.ThemeArchive.default_template_slide_node_reference",
                )?;
                selected_nested_bytes = selected_nested_bytes
                    .checked_add(payload.len())
                    .ok_or_else(DecodeError::projection)?;
                default_template_slide_node_reference =
                    Some(preflight_reference(payload, 2, budget)?);
            },
            THEME_DEFAULT_TEMPLATE_GUESS_FIELD => {
                mark_singular(
                    &mut seen,
                    6,
                    "KN.ThemeArchive.default_template_slide_node_is_our_best_guess",
                )?;
                require_varint(
                    &field,
                    "KN.ThemeArchive.default_template_slide_node_is_our_best_guess",
                )?;
                require_canonical(
                    &field,
                    "KN.ThemeArchive.default_template_slide_node_is_our_best_guess",
                )?;
                let value = field.varint.unwrap_or_default();
                if value > 1 {
                    return Err(DecodeError::noncanonical(
                        "Keynote theme bool scalar is not zero or one",
                    ));
                }
                default_template_slide_node_is_our_best_guess = Some(value == 1);
            },
            THEME_LIVE_VIDEO_FIELD => {
                mark_singular(&mut seen, 8, "KN.ThemeArchive.live_video_source_collection")?;
                let payload = selected_reference_payload(
                    &field,
                    "KN.ThemeArchive.live_video_source_collection",
                )?;
                selected_nested_bytes = selected_nested_bytes
                    .checked_add(payload.len())
                    .ok_or_else(DecodeError::projection)?;
                live_video_source_collection = Some(preflight_reference(payload, 2, budget)?);
            },
            // Fields four, eight, and ten are intentionally opaque. Their
            // outer framing is checked by `next_field`; nested schemas belong
            // to the format owners that consume those source spans.
            _ => {},
        }
    }
    Ok(Preflight {
        super_bytes: super_bytes.ok_or_else(|| DecodeError::missing("KN.ThemeArchive.super"))?,
        uuid,
        template_references: templates,
        template_payload_bytes,
        selected_nested_bytes,
        default_template_slide_node,
        default_template_slide_node_reference,
        default_template_slide_node_is_our_best_guess,
        live_video_source_collection,
    })
}

fn preflight_reference(
    source: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<ThemeReferenceSnapshot, DecodeError> {
    budget.charge_work(source.len())?;
    let mut cursor = 0usize;
    let mut seen = 0u16;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    while let Some(field) = next_field(source, &mut cursor, depth, budget)? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                mark_singular(&mut seen, 0, "TSP.Reference.identifier")?;
                require_varint(&field, "TSP.Reference.identifier")?;
                require_canonical(&field, "TSP.Reference.identifier")?;
                identifier = field.varint;
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                mark_singular(&mut seen, 1, "TSP.Reference.deprecated_type")?;
                require_varint(&field, "TSP.Reference.deprecated_type")?;
                require_canonical(&field, "TSP.Reference.deprecated_type")?;
                let value = field.varint.unwrap_or_default();
                if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
                    return Err(DecodeError::noncanonical(
                        "TSP.Reference.deprecated_type is not a sign-extended int32",
                    ));
                }
                deprecated_type = Some(decode_int32(value));
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                mark_singular(&mut seen, 2, "TSP.Reference.deprecated_is_external")?;
                require_varint(&field, "TSP.Reference.deprecated_is_external")?;
                require_canonical(&field, "TSP.Reference.deprecated_is_external")?;
                let value = field.varint.unwrap_or_default();
                if value > 1 {
                    return Err(DecodeError::noncanonical(
                        "TSP.Reference bool scalar is not zero or one",
                    ));
                }
                deprecated_is_external = Some(value == 1);
            },
            _ => {},
        }
    }
    Ok(ThemeReferenceSnapshot {
        identifier: {
            let identifier =
                identifier.ok_or_else(|| DecodeError::missing("TSP.Reference.identifier"))?;
            if identifier == 0 {
                return Err(DecodeError::invalid(
                    "TSP.Reference.identifier must be non-zero",
                ));
            }
            identifier
        },
        deprecated_type,
        deprecated_is_external,
    })
}

fn force_projection(
    source: &[u8],
    preflight: &Preflight<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    check_limit(DecodeLimit::Allocations, 0, options.max_allocations)?;
    check_limit(
        DecodeLimit::RetainedBytes,
        source.len(),
        options.max_retained_bytes,
    )?;
    check_limit(DecodeLimit::ScratchBytes, 0, options.max_scratch_bytes)?;
    let view: projection::KeynoteThemeArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    if !view.has_super() || view.super_ != preflight.super_bytes || view.uuid != preflight.uuid {
        return Err(DecodeError::projection());
    }
    let default_template_slide_node = view
        .default_template_slide_node
        .get()?
        .as_ref()
        .map(force_reference_projection)
        .transpose()?;
    let default_template_slide_node_reference = view
        .default_template_slide_node_reference
        .get()?
        .as_ref()
        .map(force_reference_projection)
        .transpose()?;
    let live_video_source_collection = view
        .live_video_source_collection
        .get()?
        .as_ref()
        .map(force_reference_projection)
        .transpose()?;
    if default_template_slide_node != preflight.default_template_slide_node
        || default_template_slide_node_reference != preflight.default_template_slide_node_reference
        || view.default_template_slide_node_is_our_best_guess
            != preflight.default_template_slide_node_is_our_best_guess
        || live_video_source_collection != preflight.live_video_source_collection
    {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn force_reference_projection(
    view: &projection::ReferenceLazyView<'_>,
) -> Result<ThemeReferenceSnapshot, DecodeError> {
    if !view.has_identifier() {
        return Err(DecodeError::missing("TSP.Reference.identifier"));
    }
    Ok(ThemeReferenceSnapshot {
        identifier: view.identifier,
        deprecated_type: view.deprecated_type,
        deprecated_is_external: view.deprecated_is_external,
    })
}

fn validate_write(
    write: &ThemeArchiveWrite<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if let Patch::Replace(payload) = write.super_bytes {
        check_limit(
            DecodeLimit::OutputBytes,
            payload.len(),
            options.max_output_bytes,
        )?;
    }
    if let Some(templates) = write.templates {
        validate_reference_slice(templates, options)?;
    }
    for reference in [
        write.default_template_slide_node,
        write.default_template_slide_node_reference,
        write.live_video_source_collection,
    ] {
        if let Patch::Replace(reference) = reference {
            validate_reference_identifier(reference)?;
        }
    }
    if (write
        .templates
        .is_some_and(|templates| !templates.is_empty())
        || matches!(write.default_template_slide_node, Patch::Replace(_))
        || matches!(
            write.default_template_slide_node_reference,
            Patch::Replace(_)
        )
        || matches!(write.live_video_source_collection, Patch::Replace(_)))
        && options.recursion_limit < 2
    {
        return Err(DecodeError::resource(
            DecodeLimit::Nesting,
            2,
            options.recursion_limit as usize,
        ));
    }
    Ok(())
}

fn validate_reference_slice(
    references: &[ThemeReferenceSnapshot],
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    for reference in references {
        validate_reference_identifier(*reference)?;
    }
    check_limit(
        DecodeLimit::TemplateReferences,
        references.len(),
        options.max_template_references,
    )?;
    // Each template emits an outer length-delimited field and at least its
    // required identifier field (four encoded bytes). Construction and
    // candidate replay visit at least eight fields and consume at least 22
    // work bytes per template. Admit this width before
    // measurement traverses caller-owned records, even when the independent
    // template-count ceiling is generous.
    let minimum_fields = references.len().checked_mul(8).ok_or_else(|| {
        DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
    })?;
    check_limit(DecodeLimit::Fields, minimum_fields, options.max_fields)?;
    let minimum_bytes = references.len().checked_mul(4).ok_or_else(|| {
        DecodeError::resource(
            DecodeLimit::OutputBytes,
            usize::MAX,
            options.max_output_bytes,
        )
    })?;
    check_limit(
        DecodeLimit::OutputBytes,
        minimum_bytes,
        options.max_output_bytes,
    )?;
    check_limit(
        DecodeLimit::RetainedBytes,
        minimum_bytes,
        options.max_retained_bytes,
    )?;
    check_limit(
        DecodeLimit::ScratchBytes,
        minimum_bytes,
        options.max_scratch_bytes,
    )?;
    let minimum_work = references.len().checked_mul(22).ok_or_else(|| {
        DecodeError::resource(DecodeLimit::WorkBytes, usize::MAX, options.max_work_bytes)
    })?;
    check_limit(DecodeLimit::WorkBytes, minimum_work, options.max_work_bytes)?;
    if !references.is_empty() {
        check_limit(DecodeLimit::Nesting, 2, options.recursion_limit as usize)?;
    }
    Ok(())
}

fn validate_reference_identifier(reference: ThemeReferenceSnapshot) -> Result<(), DecodeError> {
    if reference.identifier == 0 {
        return Err(DecodeError::invalid(
            "TSP.Reference.identifier must be non-zero",
        ));
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct RewritePlan {
    output_len: usize,
    source_fields: usize,
    output_fields: usize,
    source_template_payload_bytes: usize,
    output_template_payload_bytes: usize,
    output_selected_nested_bytes: usize,
}

fn measure_rewrite(
    source: &[u8],
    write: ThemeArchiveWrite<'_>,
    options: DecodeOptions,
) -> Result<RewritePlan, DecodeError> {
    let mut budget = Budget::new(options);
    budget.charge_work(source.len())?;
    let mut cursor = 0usize;
    let mut output_len = 0usize;
    let mut source_fields = 0usize;
    let mut output_fields = 0usize;
    let mut source_template_payload_bytes = 0usize;
    let mut output_template_payload_bytes = 0usize;
    let mut output_selected_nested_bytes = 0usize;
    let mut seen_templates = false;
    let mut seen = [false; 6];
    loop {
        let fields_before = budget.fields;
        let Some(field) = next_field(source, &mut cursor, 1, &mut budget)? else {
            break;
        };
        let mut field_cost = budget.fields - fields_before;
        if matches!(
            field.number,
            THEME_TEMPLATES_FIELD
                | THEME_DEFAULT_TEMPLATE_FIELD
                | THEME_DEFAULT_TEMPLATE_REFERENCE_FIELD
                | THEME_LIVE_VIDEO_FIELD
        ) {
            if let Some(payload) = field.payload {
                field_cost = field_cost
                    .checked_add(reference_field_count_from_payload(payload)?)
                    .ok_or_else(|| {
                        DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
                    })?;
            }
        }
        source_fields = source_fields.checked_add(field_cost).ok_or_else(|| {
            DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
        })?;
        if field.number == THEME_TEMPLATES_FIELD {
            let payload = field.payload.unwrap_or_default();
            source_template_payload_bytes = source_template_payload_bytes
                .checked_add(payload.len())
                .ok_or_else(|| {
                    DecodeError::resource(
                        DecodeLimit::WorkBytes,
                        usize::MAX,
                        options.max_work_bytes,
                    )
                })?;
            if let Some(templates) = write.templates {
                if seen_templates {
                    continue;
                }
                seen_templates = true;
                for (index, reference) in templates.iter().copied().enumerate() {
                    let raw =
                        raw_template_for_desired(source, templates, index, Some(&mut budget))?;
                    let (field_len, payload_len, nested_fields) =
                        template_emission_metrics(raw, reference)?;
                    output_len = checked_add(output_len, field_len, options.max_output_bytes)?;
                    output_fields =
                        output_fields
                            .checked_add(1 + nested_fields)
                            .ok_or_else(|| {
                                DecodeError::resource(
                                    DecodeLimit::Fields,
                                    usize::MAX,
                                    options.max_fields,
                                )
                            })?;
                    output_template_payload_bytes = output_template_payload_bytes
                        .checked_add(payload_len)
                        .ok_or_else(|| {
                            DecodeError::resource(
                                DecodeLimit::WorkBytes,
                                usize::MAX,
                                options.max_work_bytes,
                            )
                        })?;
                }
                continue;
            }
            output_len = checked_add(output_len, field.raw.len(), options.max_output_bytes)?;
            output_fields = output_fields
                .checked_add(1 + reference_field_count_from_payload(payload)?)
                .ok_or_else(|| {
                    DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
                })?;
            output_template_payload_bytes = output_template_payload_bytes
                .checked_add(payload.len())
                .ok_or_else(|| {
                    DecodeError::resource(
                        DecodeLimit::WorkBytes,
                        usize::MAX,
                        options.max_work_bytes,
                    )
                })?;
            continue;
        }
        let replacement = replacement_for_field(write, field.number);
        if let Some(replacement) = replacement {
            let index = selected_index(field.number).unwrap_or(0);
            if seen[index] {
                continue;
            }
            seen[index] = true;
            match replacement {
                Replacement::Bytes(patch) => match patch {
                    Patch::Preserve => {
                        output_len =
                            checked_add(output_len, field.raw.len(), options.max_output_bytes)?;
                        output_fields = output_fields.checked_add(field_cost).ok_or_else(|| {
                            DecodeError::resource(
                                DecodeLimit::Fields,
                                usize::MAX,
                                options.max_fields,
                            )
                        })?;
                    },
                    Patch::Clear => {},
                    Patch::Replace(payload) => {
                        output_len = checked_add(
                            output_len,
                            encoded_field_len(field.number, payload.len())?,
                            options.max_output_bytes,
                        )?;
                        output_fields = output_fields.checked_add(1).ok_or_else(|| {
                            DecodeError::resource(
                                DecodeLimit::Fields,
                                usize::MAX,
                                options.max_fields,
                            )
                        })?;
                    },
                },
                Replacement::Reference(patch) => match patch {
                    Patch::Preserve => {
                        output_len =
                            checked_add(output_len, field.raw.len(), options.max_output_bytes)?;
                        let payload = field.payload.unwrap_or_default();
                        output_fields = output_fields
                            .checked_add(1 + reference_field_count_from_payload(payload)?)
                            .ok_or_else(|| {
                                DecodeError::resource(
                                    DecodeLimit::Fields,
                                    usize::MAX,
                                    options.max_fields,
                                )
                            })?;
                        output_selected_nested_bytes = output_selected_nested_bytes
                            .checked_add(payload.len())
                            .ok_or_else(|| {
                                DecodeError::resource(
                                    DecodeLimit::WorkBytes,
                                    usize::MAX,
                                    options.max_work_bytes,
                                )
                            })?;
                    },
                    Patch::Clear => {},
                    Patch::Replace(reference) => {
                        output_len = checked_add(
                            output_len,
                            encoded_reference_field_len(field.number, reference),
                            options.max_output_bytes,
                        )?;
                        output_fields = output_fields
                            .checked_add(1 + reference_field_count(reference))
                            .ok_or_else(|| {
                                DecodeError::resource(
                                    DecodeLimit::Fields,
                                    usize::MAX,
                                    options.max_fields,
                                )
                            })?;
                        output_selected_nested_bytes = output_selected_nested_bytes
                            .checked_add(encoded_reference_payload_len(reference))
                            .ok_or_else(|| {
                                DecodeError::resource(
                                    DecodeLimit::WorkBytes,
                                    usize::MAX,
                                    options.max_work_bytes,
                                )
                            })?;
                    },
                },
                Replacement::Bool(patch) => match patch {
                    Patch::Preserve => {
                        output_len =
                            checked_add(output_len, field.raw.len(), options.max_output_bytes)?;
                        output_fields = output_fields.checked_add(field_cost).ok_or_else(|| {
                            DecodeError::resource(
                                DecodeLimit::Fields,
                                usize::MAX,
                                options.max_fields,
                            )
                        })?;
                    },
                    Patch::Clear => {},
                    Patch::Replace(_) => {
                        output_len = checked_add(output_len, 2, options.max_output_bytes)?;
                        output_fields = output_fields.checked_add(1).ok_or_else(|| {
                            DecodeError::resource(
                                DecodeLimit::Fields,
                                usize::MAX,
                                options.max_fields,
                            )
                        })?;
                    },
                },
            }
            continue;
        }
        output_len = checked_add(output_len, field.raw.len(), options.max_output_bytes)?;
        output_fields = output_fields.checked_add(field_cost).ok_or_else(|| {
            DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
        })?;
    }
    if let Some(templates) = write.templates {
        if !seen_templates {
            for (index, reference) in templates.iter().copied().enumerate() {
                let raw = raw_template_for_desired(source, templates, index, Some(&mut budget))?;
                let (field_len, payload_len, nested_fields) =
                    template_emission_metrics(raw, reference)?;
                output_len = checked_add(output_len, field_len, options.max_output_bytes)?;
                output_fields = output_fields
                    .checked_add(1 + nested_fields)
                    .ok_or_else(|| {
                        DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
                    })?;
                output_template_payload_bytes = output_template_payload_bytes
                    .checked_add(payload_len)
                    .ok_or_else(|| {
                        DecodeError::resource(
                            DecodeLimit::WorkBytes,
                            usize::MAX,
                            options.max_work_bytes,
                        )
                    })?;
            }
        }
    }
    for (field_number, replacement) in absent_replacements(write, &seen) {
        match replacement {
            Replacement::Bytes(Patch::Replace(payload)) => {
                output_len = checked_add(
                    output_len,
                    encoded_field_len(field_number, payload.len())?,
                    options.max_output_bytes,
                )?;
                output_fields = output_fields.checked_add(1).ok_or_else(|| {
                    DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
                })?;
            },
            Replacement::Reference(Patch::Replace(reference)) => {
                output_len = checked_add(
                    output_len,
                    encoded_reference_field_len(field_number, reference),
                    options.max_output_bytes,
                )?;
                output_fields = output_fields
                    .checked_add(1 + reference_field_count(reference))
                    .ok_or_else(|| {
                        DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
                    })?;
                output_selected_nested_bytes = output_selected_nested_bytes
                    .checked_add(encoded_reference_payload_len(reference))
                    .ok_or_else(|| {
                        DecodeError::resource(
                            DecodeLimit::WorkBytes,
                            usize::MAX,
                            options.max_work_bytes,
                        )
                    })?;
            },
            Replacement::Bool(Patch::Replace(_)) => {
                output_len = checked_add(output_len, 2, options.max_output_bytes)?;
                output_fields = output_fields.checked_add(1).ok_or_else(|| {
                    DecodeError::resource(DecodeLimit::Fields, usize::MAX, options.max_fields)
                })?;
            },
            _ => {},
        }
    }
    Ok(RewritePlan {
        output_len,
        source_fields,
        output_fields,
        source_template_payload_bytes,
        output_template_payload_bytes,
        output_selected_nested_bytes,
    })
}

#[derive(Debug, Clone, Copy)]
enum Replacement<'source> {
    Bytes(Patch<&'source [u8]>),
    Reference(Patch<ThemeReferenceSnapshot>),
    Bool(Patch<bool>),
}

fn replacement_for_field(
    write: ThemeArchiveWrite<'_>,
    field_number: u32,
) -> Option<Replacement<'_>> {
    match field_number {
        THEME_SUPER_FIELD => Some(Replacement::Bytes(write.super_bytes)),
        THEME_UUID_FIELD => Some(Replacement::Bytes(match write.uuid {
            Patch::Preserve => Patch::Preserve,
            Patch::Clear => Patch::Clear,
            Patch::Replace(value) => Patch::Replace(value.as_bytes()),
        })),
        THEME_DEFAULT_TEMPLATE_FIELD => {
            Some(Replacement::Reference(write.default_template_slide_node))
        },
        THEME_DEFAULT_TEMPLATE_REFERENCE_FIELD => Some(Replacement::Reference(
            write.default_template_slide_node_reference,
        )),
        THEME_DEFAULT_TEMPLATE_GUESS_FIELD => Some(Replacement::Bool(
            write.default_template_slide_node_is_our_best_guess,
        )),
        THEME_LIVE_VIDEO_FIELD => Some(Replacement::Reference(write.live_video_source_collection)),
        _ => None,
    }
}

fn absent_replacements<'source>(
    write: ThemeArchiveWrite<'source>,
    seen: &[bool; 6],
) -> [(u32, Replacement<'source>); 6] {
    let mut replacements = [
        (THEME_SUPER_FIELD, Replacement::Bytes(write.super_bytes)),
        (
            THEME_UUID_FIELD,
            Replacement::Bytes(match write.uuid {
                Patch::Preserve => Patch::Preserve,
                Patch::Clear => Patch::Clear,
                Patch::Replace(value) => Patch::Replace(value.as_bytes()),
            }),
        ),
        (
            THEME_DEFAULT_TEMPLATE_FIELD,
            Replacement::Reference(write.default_template_slide_node),
        ),
        (
            THEME_DEFAULT_TEMPLATE_REFERENCE_FIELD,
            Replacement::Reference(write.default_template_slide_node_reference),
        ),
        (
            THEME_DEFAULT_TEMPLATE_GUESS_FIELD,
            Replacement::Bool(write.default_template_slide_node_is_our_best_guess),
        ),
        (
            THEME_LIVE_VIDEO_FIELD,
            Replacement::Reference(write.live_video_source_collection),
        ),
    ];
    for (index, (_, replacement)) in replacements.iter_mut().enumerate() {
        if seen[index] {
            *replacement = match *replacement {
                Replacement::Bytes(_) => Replacement::Bytes(Patch::Preserve),
                Replacement::Reference(_) => Replacement::Reference(Patch::Preserve),
                Replacement::Bool(_) => Replacement::Bool(Patch::Preserve),
            };
        }
    }
    replacements
}

/// Return the source template record corresponding to one desired semantic
/// reference and occurrence. Matching by semantic value keeps unknown nested
/// fields and their original scalar encodings source-owned during a semantic
/// no-op rewrite. A desired record that has no source counterpart is emitted
/// canonically by the caller.
fn raw_template_for_desired<'source>(
    source: &'source [u8],
    templates: &[ThemeReferenceSnapshot],
    desired_index: usize,
    mut budget: Option<&mut Budget>,
) -> Result<Option<&'source [u8]>, DecodeError> {
    let desired = templates[desired_index];
    let desired_occurrence = templates[..desired_index]
        .iter()
        .filter(|reference| **reference == desired)
        .count();
    let mut occurrence = 0usize;
    let mut cursor = 0usize;
    let mut unmetered = Budget::new(DecodeOptions::internal_for(source.len()));
    while cursor < source.len() {
        let field = match budget.as_deref_mut() {
            Some(budget) => next_field(source, &mut cursor, 1, budget)?,
            None => next_field(source, &mut cursor, 1, &mut unmetered)?,
        };
        let Some(field) = field else { break };
        if let Some(budget) = budget.as_deref_mut() {
            budget.charge_work(field.raw.len())?;
        }
        if field.number != THEME_TEMPLATES_FIELD {
            continue;
        }
        let Some(payload) = field.payload else {
            continue;
        };
        if reference_snapshot_unchecked(payload) != Some(desired) {
            continue;
        }
        if occurrence == desired_occurrence {
            return Ok(Some(field.raw));
        }
        occurrence = occurrence
            .checked_add(1)
            .ok_or_else(|| DecodeError::resource(DecodeLimit::Fields, usize::MAX, usize::MAX))?;
    }
    Ok(None)
}

fn template_emission_metrics(
    raw: Option<&[u8]>,
    reference: ThemeReferenceSnapshot,
) -> Result<(usize, usize, usize), DecodeError> {
    if let Some(raw) = raw {
        let mut cursor = 0usize;
        let field = next_field_unchecked(raw, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("template source changed after scan"))?;
        let payload = field.payload.unwrap_or_default();
        return Ok((
            raw.len(),
            payload.len(),
            reference_field_count_from_payload(payload)?,
        ));
    }
    Ok((
        encoded_reference_field_len(THEME_TEMPLATES_FIELD, reference),
        encoded_reference_payload_len(reference),
        reference_field_count(reference),
    ))
}

fn emit_rewrite(
    source: &[u8],
    write: ThemeArchiveWrite<'_>,
    output: &mut Vec<u8>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let mut cursor = 0usize;
    let mut seen_templates = false;
    let mut seen = [false; 6];
    let mut budget = Budget::new(options);
    while let Some(field) = next_field(source, &mut cursor, 1, &mut budget)? {
        if field.number == THEME_TEMPLATES_FIELD {
            if let Some(templates) = write.templates {
                if seen_templates {
                    continue;
                }
                seen_templates = true;
                for (index, reference) in templates.iter().copied().enumerate() {
                    if let Some(raw) = raw_template_for_desired(source, templates, index, None)? {
                        output.extend_from_slice(raw);
                    } else {
                        append_reference_field(output, THEME_TEMPLATES_FIELD, reference);
                    }
                }
                continue;
            }
        }
        if let Some(replacement) = replacement_for_field(write, field.number) {
            let index = selected_index(field.number).unwrap_or(0);
            if seen[index] {
                continue;
            }
            seen[index] = true;
            append_replacement(output, field.number, replacement, field.raw)?;
        } else {
            output.extend_from_slice(field.raw);
        }
    }
    if let Some(templates) = write.templates {
        if !seen_templates {
            for (index, reference) in templates.iter().copied().enumerate() {
                if let Some(raw) = raw_template_for_desired(source, templates, index, None)? {
                    output.extend_from_slice(raw);
                } else {
                    append_reference_field(output, THEME_TEMPLATES_FIELD, reference);
                }
            }
        }
    }
    for (field_number, replacement) in absent_replacements(write, &seen) {
        append_absent_replacement(output, field_number, replacement);
    }
    Ok(())
}

fn append_replacement(
    output: &mut Vec<u8>,
    field_number: u32,
    replacement: Replacement<'_>,
    raw: &[u8],
) -> Result<(), DecodeError> {
    match replacement {
        Replacement::Bytes(Patch::Preserve) => output.extend_from_slice(raw),
        Replacement::Bytes(Patch::Clear) => {},
        Replacement::Bytes(Patch::Replace(payload)) => {
            append_length_delimited(output, field_number, payload)
        },
        Replacement::Reference(Patch::Preserve) => output.extend_from_slice(raw),
        Replacement::Reference(Patch::Clear) => {},
        Replacement::Reference(Patch::Replace(reference)) => {
            append_reference_field(output, field_number, reference)
        },
        Replacement::Bool(Patch::Preserve) => output.extend_from_slice(raw),
        Replacement::Bool(Patch::Clear) => {},
        Replacement::Bool(Patch::Replace(value)) => {
            append_varint_field(output, field_number, u64::from(value))
        },
    }
    Ok(())
}

fn append_absent_replacement(
    output: &mut Vec<u8>,
    field_number: u32,
    replacement: Replacement<'_>,
) {
    match replacement {
        Replacement::Bytes(Patch::Replace(payload)) => {
            append_length_delimited(output, field_number, payload)
        },
        Replacement::Reference(Patch::Replace(reference)) => {
            append_reference_field(output, field_number, reference)
        },
        Replacement::Bool(Patch::Replace(value)) => {
            append_varint_field(output, field_number, u64::from(value))
        },
        _ => {},
    }
}

fn same_selected_fields(
    before: ThemeArchiveSnapshot<'_>,
    after: ThemeArchiveSnapshot<'_>,
    write: ThemeArchiveWrite<'_>,
) -> bool {
    if !same_bytes_patch(before.super_bytes(), after.super_bytes(), write.super_bytes)
        || !same_string_patch(before.uuid(), after.uuid(), write.uuid)
        || !same_reference_patch(
            before.default_template_slide_node(),
            after.default_template_slide_node(),
            write.default_template_slide_node,
        )
        || !same_reference_patch(
            before.default_template_slide_node_reference(),
            after.default_template_slide_node_reference(),
            write.default_template_slide_node_reference,
        )
        || !same_bool_patch(
            before.default_template_slide_node_is_our_best_guess(),
            after.default_template_slide_node_is_our_best_guess(),
            write.default_template_slide_node_is_our_best_guess,
        )
        || !same_reference_patch(
            before.live_video_source_collection(),
            after.live_video_source_collection(),
            write.live_video_source_collection,
        )
    {
        return false;
    }
    if let Some(expected) = write.templates {
        let mut actual = after.templates();
        for reference in expected {
            if actual.next() != Some(*reference) {
                return false;
            }
        }
        actual.next().is_none()
    } else {
        let mut left = before.templates();
        let mut right = after.templates();
        loop {
            let left_value = left.next();
            let right_value = right.next();
            if left_value != right_value {
                return false;
            }
            if left_value.is_none() {
                return true;
            }
        }
    }
}

fn same_bytes_patch(before: &[u8], after: &[u8], patch: Patch<&[u8]>) -> bool {
    match patch {
        Patch::Preserve => before == after,
        Patch::Clear => false,
        Patch::Replace(value) => after == value,
    }
}

fn same_string_patch(before: Option<&str>, after: Option<&str>, patch: Patch<&str>) -> bool {
    match patch {
        Patch::Preserve => before == after,
        Patch::Clear => after.is_none(),
        Patch::Replace(value) => after == Some(value),
    }
}

fn same_reference_patch(
    before: Option<ThemeReferenceSnapshot>,
    after: Option<ThemeReferenceSnapshot>,
    patch: Patch<ThemeReferenceSnapshot>,
) -> bool {
    match patch {
        Patch::Preserve => before == after,
        Patch::Clear => after.is_none(),
        Patch::Replace(value) => after == Some(value),
    }
}

fn same_bool_patch(before: Option<bool>, after: Option<bool>, patch: Patch<bool>) -> bool {
    match patch {
        Patch::Preserve => before == after,
        Patch::Clear => after.is_none(),
        Patch::Replace(value) => after == Some(value),
    }
}

fn selected_index(field_number: u32) -> Option<usize> {
    match field_number {
        THEME_SUPER_FIELD => Some(0),
        THEME_UUID_FIELD => Some(1),
        THEME_DEFAULT_TEMPLATE_FIELD => Some(2),
        THEME_DEFAULT_TEMPLATE_REFERENCE_FIELD => Some(3),
        THEME_DEFAULT_TEMPLATE_GUESS_FIELD => Some(4),
        THEME_LIVE_VIDEO_FIELD => Some(5),
        _ => None,
    }
}

fn ensure_requirements(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    check_limit(
        DecodeLimit::OutputBytes,
        requirements.output_bytes,
        limits.output_bytes,
    )?;
    check_limit(DecodeLimit::Fields, requirements.fields, limits.fields)?;
    check_limit(
        DecodeLimit::WorkBytes,
        requirements.work_bytes,
        limits.work_bytes,
    )?;
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::resource(
            DecodeLimit::Nesting,
            requirements.max_depth as usize,
            limits.max_depth as usize,
        ));
    }
    check_limit(
        DecodeLimit::Allocations,
        requirements.allocations,
        limits.allocations,
    )?;
    check_limit(
        DecodeLimit::RetainedBytes,
        requirements.retained_bytes,
        limits.retained_bytes,
    )?;
    check_limit(
        DecodeLimit::ScratchBytes,
        requirements.scratch_bytes,
        limits.scratch_bytes,
    )
}

fn admit_output(
    options: &DecodeOptions,
    output_bytes: usize,
    fields: usize,
    work: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
) -> Result<(), DecodeError> {
    check_limit(
        DecodeLimit::OutputBytes,
        output_bytes,
        options.max_output_bytes,
    )?;
    check_limit(DecodeLimit::Fields, fields, options.max_fields)?;
    check_limit(DecodeLimit::WorkBytes, work, options.max_work_bytes)?;
    check_limit(
        DecodeLimit::Allocations,
        allocations,
        options.max_allocations,
    )?;
    check_limit(
        DecodeLimit::RetainedBytes,
        retained,
        options.max_retained_bytes,
    )?;
    check_limit(
        DecodeLimit::ScratchBytes,
        scratch,
        options.max_scratch_bytes,
    )
}

fn check_limit(limit: DecodeLimit, observed: usize, maximum: usize) -> Result<(), DecodeError> {
    if observed > maximum {
        return Err(DecodeError::resource(limit, observed, maximum));
    }
    Ok(())
}

fn ensure_input(options: &DecodeOptions, input_bytes: usize) -> Result<(), DecodeError> {
    check_limit(
        DecodeLimit::InputBytes,
        input_bytes,
        options.max_message_bytes,
    )
}

fn validate_options(options: DecodeOptions) -> Result<(), DecodeError> {
    if options.max_message_bytes > MAX_AUTOMATIC_LIMIT {
        return Err(DecodeError::resource(
            DecodeLimit::InputBytes,
            options.max_message_bytes,
            MAX_AUTOMATIC_LIMIT,
        ));
    }
    for (limit, value) in [
        (DecodeLimit::Fields, options.max_fields),
        (DecodeLimit::WorkBytes, options.max_work_bytes),
        (DecodeLimit::OutputBytes, options.max_output_bytes),
        (DecodeLimit::Allocations, options.max_allocations),
        (DecodeLimit::RetainedBytes, options.max_retained_bytes),
        (DecodeLimit::ScratchBytes, options.max_scratch_bytes),
        (
            DecodeLimit::TemplateReferences,
            options.max_template_references,
        ),
    ] {
        if value > MAX_AUTOMATIC_LIMIT {
            return Err(DecodeError::resource(limit, value, MAX_AUTOMATIC_LIMIT));
        }
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::resource(
            DecodeLimit::Nesting,
            options.recursion_limit as usize,
            MAX_RECURSION_LIMIT as usize,
        ));
    }
    Ok(())
}

fn checked_add(left: usize, right: usize, maximum: usize) -> Result<usize, DecodeError> {
    let value = left
        .checked_add(right)
        .ok_or_else(|| DecodeError::resource(DecodeLimit::OutputBytes, usize::MAX, maximum))?;
    check_limit(DecodeLimit::OutputBytes, value, maximum)?;
    Ok(value)
}

fn mark_singular(seen: &mut u16, index: u32, name: &'static str) -> Result<(), DecodeError> {
    let bit = 1u16
        .checked_shl(index)
        .ok_or_else(|| DecodeError::noncanonical("selected field number exceeds presence mask"))?;
    if *seen & bit != 0 {
        return Err(DecodeError::duplicate(name));
    }
    *seen |= bit;
    Ok(())
}

fn require_length_delimited(field: &WireField<'_>, name: &'static str) -> Result<(), DecodeError> {
    if field.wire_type != 2 {
        return Err(DecodeError::invalid(format!(
            "{name} has the wrong wire type"
        )));
    }
    Ok(())
}

fn require_varint(field: &WireField<'_>, name: &'static str) -> Result<(), DecodeError> {
    if field.wire_type != 0 {
        return Err(DecodeError::invalid(format!(
            "{name} has the wrong wire type"
        )));
    }
    Ok(())
}

fn require_canonical(field: &WireField<'_>, name: &'static str) -> Result<(), DecodeError> {
    if !field.tag_canonical || !field.value_canonical {
        return Err(DecodeError::noncanonical(name));
    }
    Ok(())
}

fn selected_reference_payload<'source>(
    field: &WireField<'source>,
    name: &'static str,
) -> Result<&'source [u8], DecodeError> {
    require_length_delimited(field, name)?;
    require_canonical(field, name)?;
    Ok(field.payload.unwrap_or_default())
}

fn decode_int32(value: u64) -> i32 {
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        reason = "The selected int32 field is checked for canonical protobuf representation."
    )]
    {
        value as i32
    }
}

fn next_field<'source>(
    source: &'source [u8],
    cursor: &mut usize,
    depth: u32,
    budget: &mut Budget,
) -> Result<Option<WireField<'source>>, DecodeError> {
    if *cursor == source.len() {
        return Ok(None);
    }
    if depth == 0 || depth > budget.options.recursion_limit {
        return Err(DecodeError::resource(
            DecodeLimit::Nesting,
            depth as usize,
            budget.options.recursion_limit as usize,
        ));
    }
    let start = *cursor;
    let (tag, tag_canonical) = read_varint(source, cursor)?;
    let field_number = u32::try_from(tag >> 3)
        .map_err(|_| DecodeError::invalid("protobuf field number overflows u32"))?;
    let wire_type =
        u8::try_from(tag & 7).map_err(|_| DecodeError::invalid("invalid protobuf wire type"))?;
    if field_number == 0 || field_number > 0x1fff_ffff {
        return Err(DecodeError::invalid("protobuf field number is invalid"));
    }
    budget.field(depth)?;
    let mut payload = None;
    let mut varint = None;
    let mut value_canonical = true;
    match wire_type {
        0 => {
            let (value, canonical) = read_varint(source, cursor)?;
            varint = Some(value);
            value_canonical = canonical;
        },
        1 => {
            take_exact(source, cursor, 8)?;
        },
        2 => {
            let (length, canonical) = read_varint(source, cursor)?;
            value_canonical = canonical;
            let length = usize::try_from(length)
                .map_err(|_| DecodeError::invalid("length-delimited field is too large"))?;
            let end = cursor
                .checked_add(length)
                .ok_or_else(|| DecodeError::invalid("length-delimited field overflows"))?;
            if end > source.len() {
                return Err(DecodeError::invalid("length-delimited field is truncated"));
            }
            payload = Some(&source[*cursor..end]);
            *cursor = end;
        },
        3 => skip_group(source, cursor, field_number, depth, budget)?,
        4 => return Err(DecodeError::invalid("unexpected protobuf end-group")),
        5 => {
            take_exact(source, cursor, 4)?;
        },
        _ => return Err(DecodeError::invalid("invalid protobuf wire type")),
    }
    Ok(Some(WireField {
        number: field_number,
        wire_type,
        payload,
        varint,
        raw: &source[start..*cursor],
        tag_canonical,
        value_canonical,
    }))
}

fn skip_group(
    source: &[u8],
    cursor: &mut usize,
    expected_field: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    if depth >= budget.options.recursion_limit {
        return Err(DecodeError::resource(
            DecodeLimit::Nesting,
            (depth + 1) as usize,
            budget.options.recursion_limit as usize,
        ));
    }
    loop {
        if *cursor == source.len() {
            return Err(DecodeError::invalid("unterminated protobuf group"));
        }
        let (tag, _) = read_varint(source, cursor)?;
        let number = u32::try_from(tag >> 3)
            .map_err(|_| DecodeError::invalid("protobuf group field number overflows u32"))?;
        let wire_type =
            u8::try_from(tag & 7).map_err(|_| DecodeError::invalid("invalid group wire type"))?;
        if number == 0 {
            return Err(DecodeError::invalid("protobuf group field number is zero"));
        }
        budget.field(depth + 1)?;
        match wire_type {
            4 => {
                if number != expected_field {
                    return Err(DecodeError::invalid(
                        "protobuf group closes the wrong field",
                    ));
                }
                return Ok(());
            },
            3 => skip_group(source, cursor, number, depth + 1, budget)?,
            0 => {
                read_varint(source, cursor)?;
            },
            1 => {
                take_exact(source, cursor, 8)?;
            },
            2 => {
                let (length, _) = read_varint(source, cursor)?;
                let length = usize::try_from(length)
                    .map_err(|_| DecodeError::invalid("group field is too large"))?;
                take_exact(source, cursor, length)?;
            },
            5 => {
                take_exact(source, cursor, 4)?;
            },
            _ => return Err(DecodeError::invalid("invalid protobuf group wire type")),
        }
    }
}

fn read_varint(source: &[u8], cursor: &mut usize) -> Result<(u64, bool), DecodeError> {
    let start = *cursor;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *source
            .get(*cursor)
            .ok_or_else(|| DecodeError::invalid("truncated protobuf varint"))?;
        *cursor += 1;
        if index == 9 && byte > 1 {
            return Err(DecodeError::invalid("protobuf varint overflows u64"));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = *cursor - start;
            return Ok((value, consumed == canonical_varint_len(value)));
        }
    }
    Err(DecodeError::invalid("protobuf varint is too long"))
}

fn canonical_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn take_exact<'source>(
    source: &'source [u8],
    cursor: &mut usize,
    length: usize,
) -> Result<&'source [u8], DecodeError> {
    let end = cursor
        .checked_add(length)
        .ok_or_else(|| DecodeError::invalid("protobuf field length overflows"))?;
    if end > source.len() {
        return Err(DecodeError::invalid(
            "protobuf fixed-width field is truncated",
        ));
    }
    let payload = &source[*cursor..end];
    *cursor = end;
    Ok(payload)
}

fn next_field_unchecked<'source>(
    source: &'source [u8],
    cursor: &mut usize,
) -> Option<WireField<'source>> {
    let mut budget = Budget::new(DecodeOptions::internal_for(source.len()));
    next_field(source, cursor, 1, &mut budget).ok().flatten()
}

fn reference_snapshot_unchecked(source: &[u8]) -> Option<ThemeReferenceSnapshot> {
    let mut cursor = 0usize;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    while let Some(field) = next_field_unchecked(source, &mut cursor) {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD if field.wire_type == 0 => identifier = field.varint,
            REFERENCE_DEPRECATED_TYPE_FIELD if field.wire_type == 0 => {
                deprecated_type = field.varint.map(decode_int32)
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD if field.wire_type == 0 => {
                deprecated_is_external = field.varint.map(|value| value == 1)
            },
            _ => {},
        }
    }
    identifier.map(|identifier| ThemeReferenceSnapshot {
        identifier,
        deprecated_type,
        deprecated_is_external,
    })
}

fn encoded_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn encoded_field_len(field_number: u32, payload_len: usize) -> Result<usize, DecodeError> {
    encoded_varint_len(u64::from(field_number) << 3 | 2)
        .checked_add(encoded_varint_len(payload_len as u64))
        .and_then(|value| value.checked_add(payload_len))
        .ok_or_else(|| {
            DecodeError::resource(DecodeLimit::OutputBytes, usize::MAX, MAX_AUTOMATIC_LIMIT)
        })
}

fn encoded_reference_payload_len(reference: ThemeReferenceSnapshot) -> usize {
    let mut length = encoded_varint_len(u64::from(REFERENCE_IDENTIFIER_FIELD) << 3)
        + encoded_varint_len(reference.identifier);
    if let Some(value) = reference.deprecated_type {
        length += encoded_varint_len(u64::from(REFERENCE_DEPRECATED_TYPE_FIELD) << 3)
            + encoded_varint_len(value as i64 as u64);
    }
    if reference.deprecated_is_external.is_some() {
        length += encoded_varint_len(u64::from(REFERENCE_DEPRECATED_EXTERNAL_FIELD) << 3) + 1;
    }
    length
}

fn reference_field_count(reference: ThemeReferenceSnapshot) -> usize {
    1usize
        + usize::from(reference.deprecated_type.is_some())
        + usize::from(reference.deprecated_is_external.is_some())
}

fn reference_field_count_from_payload(payload: &[u8]) -> Result<usize, DecodeError> {
    let mut budget = Budget::new(DecodeOptions::internal_for(payload.len()));
    let mut cursor = 0usize;
    while next_field(payload, &mut cursor, 1, &mut budget)?.is_some() {}
    Ok(budget.fields)
}

fn encoded_reference_field_len(field_number: u32, reference: ThemeReferenceSnapshot) -> usize {
    encoded_field_len(field_number, encoded_reference_payload_len(reference)).unwrap_or(usize::MAX)
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        #[allow(
            clippy::cast_possible_truncation,
            reason = "Only low seven value bits are emitted per protobuf varint byte."
        )]
        let byte = value as u8;
        output.push(byte & 0x7f | 0x80);
        value >>= 7;
    }
    #[allow(
        clippy::cast_possible_truncation,
        reason = "The final protobuf varint byte contains at most seven bits."
    )]
    output.push(value as u8);
}

fn append_varint_field(output: &mut Vec<u8>, field_number: u32, value: u64) {
    append_varint(output, u64::from(field_number) << 3);
    append_varint(output, value);
}

fn append_length_delimited(output: &mut Vec<u8>, field_number: u32, payload: &[u8]) {
    append_varint(output, u64::from(field_number) << 3 | 2);
    append_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn append_reference_field(
    output: &mut Vec<u8>,
    field_number: u32,
    reference: ThemeReferenceSnapshot,
) {
    let payload_len = encoded_reference_payload_len(reference);
    append_varint(output, u64::from(field_number) << 3 | 2);
    append_varint(output, payload_len as u64);
    append_reference_payload(output, reference);
}

fn append_reference_payload(output: &mut Vec<u8>, reference: ThemeReferenceSnapshot) {
    append_varint_field(output, REFERENCE_IDENTIFIER_FIELD, reference.identifier);
    if let Some(value) = reference.deprecated_type {
        append_varint_field(output, REFERENCE_DEPRECATED_TYPE_FIELD, value as i64 as u64);
    }
    if let Some(value) = reference.deprecated_is_external {
        append_varint_field(
            output,
            REFERENCE_DEPRECATED_EXTERNAL_FIELD,
            u64::from(value),
        );
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn varint_field(field: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        append_varint_field(&mut output, field, value);
        output
    }

    fn ref_payload(identifier: u64) -> Vec<u8> {
        varint_field(REFERENCE_IDENTIFIER_FIELD, identifier)
    }

    fn length_field(field: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        append_length_delimited(&mut output, field, payload);
        output
    }

    fn minimal() -> Vec<u8> {
        let mut source = length_field(THEME_SUPER_FIELD, &[0xa5, 0x02]);
        source.extend(length_field(THEME_TEMPLATES_FIELD, &ref_payload(4)));
        source.extend(length_field(THEME_TEMPLATES_FIELD, &ref_payload(9)));
        source.extend(length_field(THEME_UUID_FIELD, b"theme-uuid"));
        source.extend(length_field(THEME_DEFAULT_TEMPLATE_FIELD, &ref_payload(7)));
        source.extend(varint_field(THEME_DEFAULT_TEMPLATE_GUESS_FIELD, 1));
        source.extend(length_field(THEME_LIVE_VIDEO_FIELD, &ref_payload(11)));
        source
    }

    #[test]
    fn borrowed_projection_preserves_order_and_selected_values() {
        let source = minimal();
        let snapshot = decode_theme_archive(&source, DecodeOptions::for_source(&source))
            .expect("minimal outer theme should decode");
        assert_eq!(snapshot.super_bytes(), &[0xa5, 0x02]);
        assert_eq!(snapshot.uuid(), Some("theme-uuid"));
        assert_eq!(snapshot.template_count(), 2);
        assert_eq!(
            snapshot
                .templates()
                .map(|value| value.identifier)
                .collect::<Vec<_>>(),
            [4, 9]
        );
        assert_eq!(
            snapshot
                .default_template_slide_node()
                .map(|value| value.identifier),
            Some(7)
        );
        assert_eq!(
            snapshot.default_template_slide_node_is_our_best_guess(),
            Some(true)
        );
        assert_eq!(
            snapshot
                .live_video_source_collection()
                .map(|value| value.identifier),
            Some(11)
        );
    }

    #[test]
    fn selected_malformed_fields_are_rejected_and_unknown_fields_are_opaque() {
        let mut source = minimal();
        source.extend(length_field(4, &[0xff, 0x00]));
        source.extend(length_field(8, &[0xff]));
        source.extend(length_field(10, &[0xfe]));
        source.extend_from_slice(&[0xa0, 0x06, 0x80, 0x00]);
        assert!(decode_theme_archive(&source, DecodeOptions::for_source(&source)).is_ok());

        let missing = length_field(THEME_TEMPLATES_FIELD, &ref_payload(1));
        let error = decode_theme_archive(&missing, DecodeOptions::for_source(&missing))
            .expect_err("required super must be present");
        assert_eq!(
            error.missing_required_field(),
            Some("KN.ThemeArchive.super")
        );

        let mut duplicate = minimal();
        duplicate.extend(length_field(THEME_UUID_FIELD, b"duplicate"));
        let error = decode_theme_archive(&duplicate, DecodeOptions::for_source(&duplicate))
            .expect_err("uuid duplicate must be rejected");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("KN.ThemeArchive.uuid")
        );
    }

    #[test]
    fn native_reference_identifiers_are_non_zero() {
        let source = [
            length_field(THEME_SUPER_FIELD, &[0xa5]),
            length_field(THEME_TEMPLATES_FIELD, &ref_payload(0)),
        ]
        .concat();
        let error = decode_theme_archive(&source, DecodeOptions::for_source(&source))
            .expect_err("zero native identifiers must be rejected");
        assert!(matches!(
            error.kind(),
            DecodeErrorKind::InvalidFormat(message) if message.contains("non-zero")
        ));

        let zero = [ThemeReferenceSnapshot::new(0)];
        let error = ThemeArchiveBuilder::new()
            .with_super(&[1])
            .with_templates(&zero)
            .encode(DecodeOptions::default())
            .expect_err("zero builder identifiers must be rejected before allocation");
        assert!(matches!(
            error.kind(),
            DecodeErrorKind::InvalidFormat(message) if message.contains("non-zero")
        ));

        let error = ThemeArchiveBuilder::new()
            .with_super(&[1])
            .with_default_template_slide_node(ThemeReferenceSnapshot::new(0))
            .encode(DecodeOptions::default())
            .expect_err("zero optional native identifiers must be rejected");
        assert!(matches!(
            error.kind(),
            DecodeErrorKind::InvalidFormat(message) if message.contains("non-zero")
        ));
    }

    #[test]
    fn template_rewrite_preserves_unknown_nested_reference_fields() {
        let mut reference = ref_payload(1);
        reference.extend(length_field(77, b"future-template-field"));
        let source = [
            length_field(THEME_SUPER_FIELD, &[0xa5]),
            length_field(THEME_TEMPLATES_FIELD, &reference),
        ]
        .concat();
        let templates = [ThemeReferenceSnapshot::new(1)];
        let output = rewrite_theme_archive(
            &source,
            ThemeArchiveWrite::new().with_templates(&templates),
            DecodeOptions::for_source(&source),
        )
        .expect("semantic template no-op should rewrite");
        assert_eq!(output, source);
    }

    #[test]
    fn public_resource_policies_are_finite_before_scanning() {
        let source = minimal();
        let error = decode_theme_archive(
            &source,
            DecodeOptions::new(usize::MAX, usize::MAX, usize::MAX, MAX_RECURSION_LIMIT),
        )
        .expect_err("unbounded public policies must be rejected");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::InputBytes));
        let error = decode_theme_archive(
            &source,
            DecodeOptions::for_source(&source).with_max_depth(0),
        )
        .expect_err("zero nesting policies must be rejected");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::Nesting));
    }

    #[test]
    fn rewrite_copies_opaque_spans_and_changes_selected_fields() {
        let mut source = minimal();
        let unknown = length_field(4, &[0x0a, 0x01, 0xff]);
        source.extend_from_slice(&unknown);
        let templates = [
            ThemeReferenceSnapshot::new(19),
            ThemeReferenceSnapshot::new(23),
        ];
        let write = ThemeArchiveWrite::new()
            .with_uuid("replacement")
            .with_templates(&templates)
            .with_live_video_source_collection(ThemeReferenceSnapshot::new(31));
        let options = DecodeOptions::for_source(&source)
            .with_max_output_bytes(source.len() * 3)
            .with_max_retained_bytes(source.len() * 6)
            .with_max_work_bytes(source.len() * 100)
            .with_max_fields(source.len() * 30);
        let output = rewrite_theme_archive(&source, write, options).expect("rewrite should pass");
        let snapshot = decode_theme_archive(&output, DecodeOptions::for_source(&output))
            .expect("rewritten theme should decode");
        assert_eq!(snapshot.uuid(), Some("replacement"));
        assert_eq!(
            snapshot
                .templates()
                .map(|value| value.identifier)
                .collect::<Vec<_>>(),
            [19, 23]
        );
        assert_eq!(
            snapshot
                .live_video_source_collection()
                .map(|value| value.identifier),
            Some(31)
        );
        assert!(output.ends_with(&unknown));
    }

    #[test]
    fn builder_emits_native_order_and_limits_before_allocation() {
        let templates = [
            ThemeReferenceSnapshot::new(2),
            ThemeReferenceSnapshot::new(3),
        ];
        let builder = ThemeArchiveBuilder::new()
            .with_super(&[0x01, 0x02])
            .with_templates(&templates)
            .with_uuid("fresh")
            .with_default_template_slide_node(ThemeReferenceSnapshot::new(5))
            .with_default_template_slide_node_is_our_best_guess(false);
        let output = builder
            .encode(DecodeOptions::default())
            .expect("fresh builder should encode");
        let snapshot = decode_theme_archive(&output, DecodeOptions::for_source(&output))
            .expect("builder output should decode");
        assert_eq!(snapshot.super_bytes(), &[1, 2]);
        assert_eq!(snapshot.uuid(), Some("fresh"));
        assert_eq!(
            snapshot
                .templates()
                .map(|value| value.identifier)
                .collect::<Vec<_>>(),
            [2, 3]
        );
        assert_eq!(
            snapshot.default_template_slide_node_is_our_best_guess(),
            Some(false)
        );
    }

    #[test]
    fn tiny_work_and_output_ceilings_are_one_under() {
        let source = minimal();
        let (_, report) =
            decode_theme_archive_with_report(&source, DecodeOptions::for_source(&source))
                .expect("baseline report should decode");
        assert!(
            decode_theme_archive(
                &source,
                DecodeOptions::for_source(&source).with_max_work_bytes(report.work_bytes() - 1)
            )
            .is_err()
        );
        assert!(
            decode_theme_archive(
                &source,
                DecodeOptions::for_source(&source).with_max_fields(report.fields() - 1)
            )
            .is_err()
        );
        let excessive_depth = decode_theme_archive(
            &source,
            DecodeOptions::for_source(&source).with_max_depth(MAX_RECURSION_LIMIT + 1),
        )
        .expect_err("public depth above the hard ceiling must be rejected before scanning");
        assert_eq!(excessive_depth.resource_limit(), Some(DecodeLimit::Nesting));
        let builder = ThemeArchiveBuilder::new().with_super(&[1, 2, 3]);
        assert!(
            builder
                .encode(DecodeOptions::default().with_max_output_bytes(1))
                .is_err()
        );
    }

    #[test]
    fn public_depth_ceiling_precedes_malformed_group_scan() {
        let malformed_group = [0x0b, 0x0c];
        let error = decode_theme_archive(
            &malformed_group,
            DecodeOptions::for_source(&malformed_group).with_max_depth(MAX_RECURSION_LIMIT + 1),
        )
        .expect_err("the public depth ceiling must be checked before wire traversal");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::Nesting));
    }

    #[test]
    fn builder_checks_nested_depth_before_candidate_allocation() {
        let templates = [ThemeReferenceSnapshot::new(17)];
        let error = ThemeArchiveBuilder::new()
            .with_super(&[1])
            .with_templates(&templates)
            .encode(DecodeOptions::default().with_max_depth(1))
            .expect_err("a nested reference needs one level below the outer archive");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::Nesting));

        let error = ThemeArchiveBuilder::new()
            .with_super(&[1])
            .encode(DecodeOptions::default().with_max_depth(MAX_RECURSION_LIMIT + 1))
            .expect_err("public depth above the hard ceiling must be rejected");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::Nesting));
    }

    #[test]
    fn prepared_rewrite_exact_limits_include_nested_candidate_and_comparison() {
        let source = minimal();
        let templates = [
            ThemeReferenceSnapshot {
                identifier: 19,
                deprecated_type: Some(-7),
                deprecated_is_external: Some(true),
            },
            ThemeReferenceSnapshot::new(23),
        ];
        let write = ThemeArchiveWrite::new()
            .with_templates(&templates)
            .with_uuid("replacement")
            .with_live_video_source_collection(ThemeReferenceSnapshot::new(31));
        let options = DecodeOptions::for_source(&source)
            .with_max_output_bytes(source.len() * 4)
            .with_max_retained_bytes(source.len() * 8)
            .with_max_work_bytes(source.len() * 200)
            .with_max_fields(source.len() * 100);

        let prepared = prepare_theme_archive_rewrite(&source, write, options)
            .expect("rewrite requirements should be admitted before execution");
        let requirements = prepared.execution_requirements();
        assert_eq!(requirements.max_depth, 2);
        assert!(requirements.fields > 3);
        assert!(requirements.work_bytes > source.len() * 3);

        let error = prepared
            .execute(
                RewriteExecutionLimits::exact(requirements).with_fields(requirements.fields - 1),
            )
            .expect_err("one field below the candidate transaction must reject preallocation");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::Fields));

        let prepared = prepare_theme_archive_rewrite(&source, write, options)
            .expect("field retry should still prepare without allocation");
        let requirements = prepared.execution_requirements();
        let error = prepared
            .execute(
                RewriteExecutionLimits::exact(requirements)
                    .with_work_bytes(requirements.work_bytes - 1),
            )
            .expect_err("one work byte below candidate replay must reject preallocation");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::WorkBytes));

        let prepared = prepare_theme_archive_rewrite(&source, write, options)
            .expect("depth retry should still prepare without allocation");
        let requirements = prepared.execution_requirements();
        let error = prepared
            .execute(
                RewriteExecutionLimits::exact(requirements)
                    .with_max_depth(requirements.max_depth - 1),
            )
            .expect_err("candidate nested references need the admitted depth");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::Nesting));

        let prepared = prepare_theme_archive_rewrite(&source, write, options)
            .expect("exact retry should still prepare without allocation");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("exact preadmitted limits should execute");
        assert_eq!(output.report().fields(), requirements.fields);
        assert_eq!(output.report().work_bytes(), requirements.work_bytes);
    }

    #[test]
    fn template_width_is_admitted_before_measurement() {
        let templates = [ThemeReferenceSnapshot::new(1); 4];
        let options = DecodeOptions::for_source(&[0; 128]);
        for (options, expected) in [
            (options.with_max_fields(7), DecodeLimit::Fields),
            (options.with_max_work_bytes(87), DecodeLimit::WorkBytes),
            (
                options.with_max_retained_bytes(15),
                DecodeLimit::RetainedBytes,
            ),
            (
                options.with_max_scratch_bytes(15),
                DecodeLimit::ScratchBytes,
            ),
            (options.with_max_output_bytes(15), DecodeLimit::OutputBytes),
            (options.with_max_depth(1), DecodeLimit::Nesting),
        ] {
            let error = validate_reference_slice(&templates, options)
                .expect_err("width must be admitted without walking the templates");
            assert_eq!(error.resource_limit(), Some(expected));
        }
    }

    #[test]
    fn encode_requirements_are_exact_before_output_allocation() {
        let super_bytes = [0x90, 0x01];
        let templates = [ThemeReferenceSnapshot::new(12)];
        let builder = ThemeArchiveBuilder::new()
            .with_super(&super_bytes)
            .with_templates(&templates)
            .with_uuid("oracle")
            .with_default_template_slide_node(ThemeReferenceSnapshot::new(14));
        let requirements = builder
            .encode_requirements(DecodeOptions::default())
            .expect("bounded measurement");
        let expected_length = requirements.output_bytes();
        let exact_options = requirements.exact_options();
        let below_options = exact_options.with_max_output_bytes(expected_length - 1);
        let public_output = builder
            .encode(exact_options)
            .expect("public encode wrapper");
        let output = requirements.encode().expect("exact measured limits");
        assert_eq!(output.len(), expected_length);
        assert_eq!(public_output, output);

        let error = builder
            .encode_requirements(below_options)
            .expect_err("one output byte below the measured archive must fail");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::OutputBytes));
    }

    #[test]
    fn prost_schema_oracle_matches_selected_outer_fields() -> Result<(), Box<dyn std::error::Error>>
    {
        use crate::{kn, tsp, tss};
        use prost::Message as _;

        let native = kn::ThemeArchive {
            super_: tss::ThemeArchive::default(),
            templates: vec![tsp::Reference {
                identifier: 12,
                deprecated_type: Some(7),
                deprecated_is_external: Some(false),
            }],
            uuid: Some("oracle".to_owned()),
            classic_theme_records: Vec::new(),
            default_template_slide_node: Some(tsp::Reference {
                identifier: 14,
                deprecated_type: None,
                deprecated_is_external: None,
            }),
            default_template_slide_node_reference: None,
            default_template_slide_node_is_our_best_guess: Some(true),
            custom_effect_timing_curves: Vec::new(),
            live_video_source_collection: Some(tsp::Reference {
                identifier: 16,
                deprecated_type: None,
                deprecated_is_external: None,
            }),
            motion_background_style_presets: Vec::new(),
        };
        let mut source = length_field(THEME_SUPER_FIELD, &[0x90, 0x01]);
        let mut template = ref_payload(12);
        template.extend(varint_field(REFERENCE_DEPRECATED_TYPE_FIELD, 7));
        source.extend(length_field(THEME_TEMPLATES_FIELD, &template));
        source.extend(length_field(THEME_UUID_FIELD, b"oracle"));
        source.extend(length_field(THEME_DEFAULT_TEMPLATE_FIELD, &ref_payload(14)));
        source.extend(varint_field(THEME_DEFAULT_TEMPLATE_GUESS_FIELD, 1));
        source.extend(length_field(THEME_LIVE_VIDEO_FIELD, &ref_payload(16)));
        let snapshot = decode_theme_archive(&source, DecodeOptions::for_source(&source))?;
        let native_bytes = native.encode_to_vec();
        let _ = kn::ThemeArchive::decode(native_bytes.as_slice())?;
        assert_eq!(snapshot.uuid(), native.uuid.as_deref());
        assert_eq!(
            snapshot
                .templates()
                .map(|value| value.identifier)
                .collect::<Vec<_>>(),
            [12]
        );
        assert_eq!(
            snapshot
                .default_template_slide_node()
                .map(|value| value.identifier),
            Some(14)
        );
        assert_eq!(
            snapshot
                .live_video_source_collection()
                .map(|value| value.identifier),
            Some(16)
        );
        Ok(())
    }

    #[test]
    fn zero_reference_identifier_is_a_typed_error() {
        assert!(ThemeReferenceSnapshot::new(0).encode().is_err());
        assert!(!ThemeReferenceSnapshot::new(7).encode().unwrap().is_empty());
    }
}
