//! Bounded, source-preserving ownership for the shared iWork theme archive.
//!
//! TSS.ThemeArchive uses proto2 extension numbers for five application preset
//! families. A generated base message cannot be used as a round-trip owner:
//! it would either discard extension bytes or require an input-width generated
//! repeated graph. This module keeps the complete wrapper and base source
//! bytes authoritative, validates the native wire shape with finite limits,
//! and uses a private lazy Buffa view as an independent check for the
//! singular envelope and identifier fields. Format-focused builders may
//! provide complete encoded preset records for a new package; live catalog
//! edits use the bounded typed reference-list rewrite below, so the host does
//! not need to own a generated `IWorkThemeArchive` graph.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict wire pass intentionally precedes the private Buffa view."
)]

use std::{fmt, mem::size_of, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_theme_archive_generated::LitchiIwaThemeProjection as projection;

const THEME_SUPER_FIELD: u32 = 1;
const LEGACY_STYLESHEET_FIELD: u32 = 1;
const THEME_IDENTIFIER_FIELD: u32 = 3;
const DOCUMENT_STYLESHEET_FIELD: u32 = 4;
const OLD_PRESET_UUIDS_FIELD: u32 = 5;
const NEW_PRESET_UUIDS_FIELD: u32 = 6;
const COLOR_PRESETS_FIELD: u32 = 10;
const DRAWING_PRESETS_FIELD: u32 = 100;
const TEXT_PRESETS_FIELD: u32 = 110;
const CHART_PRESETS_FIELD: u32 = 120;
const TABLE_PRESETS_FIELD: u32 = 200;
const APPLICATION_PRESETS_FIELD: u32 = 210;
const MAX_AUTOMATIC_LIMIT: usize = 64 * 1024 * 1024;
const MAX_RECURSION_LIMIT: u32 = 64;

const EXTENSION_FIELDS: [u32; 5] = [
    DRAWING_PRESETS_FIELD,
    TEXT_PRESETS_FIELD,
    CHART_PRESETS_FIELD,
    TABLE_PRESETS_FIELD,
    APPLICATION_PRESETS_FIELD,
];

/// Which finite resource ceiling rejected an input or rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecodeLimit {
    /// Source bytes inspected by the transaction.
    InputBytes,
    /// Number of encoded fields visited.
    Fields,
    /// Aggregate strict-scan and Buffa work bytes.
    WorkBytes,
    /// Candidate output bytes.
    OutputBytes,
    /// Logical output allocation events.
    Allocations,
    /// Source and candidate bytes retained together.
    RetainedBytes,
    /// Candidate staging bytes.
    ScratchBytes,
    /// Maximum nested protobuf depth.
    Nesting,
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
        };
        formatter.write_str(name)
    }
}

/// Stable category for a theme decode or rewrite failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DecodeErrorKind {
    /// Malformed protobuf framing or a schema-directed field violation.
    InvalidFormat(String),
    /// A required schema field was absent.
    MissingRequired(String),
    /// A selected scalar or wire representation was not canonical.
    NonCanonical(String),
    /// A finite input, output, or aggregate budget was exceeded.
    Resource {
        /// The rejected resource category.
        limit: DecodeLimit,
        /// Amount requested or observed.
        observed: usize,
        /// Configured ceiling.
        maximum: usize,
    },
    /// A singular native field appeared more than once where the owner needs
    /// one unambiguous extension or wrapper edge.
    DuplicateSingular(String),
    /// The generated lazy projection disagreed with the strict source pass.
    Projection,
    /// The requested candidate capacity could not be reserved.
    Allocation { amount: usize },
}

/// Error returned by the bounded theme owner.
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

    fn missing(field: impl Into<String>) -> Self {
        Self {
            kind: DecodeErrorKind::MissingRequired(field.into()),
        }
    }

    fn noncanonical(reason: impl Into<String>) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason.into()),
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

    fn duplicate(field: impl Into<String>) -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateSingular(field.into()),
        }
    }

    fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    fn allocation(amount: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation { amount },
        }
    }

    /// Return the stable error category.
    #[must_use]
    pub const fn kind(&self) -> &DecodeErrorKind {
        &self.kind
    }

    /// Return the rejected resource category, if this was a finite-limit
    /// failure.
    #[must_use]
    pub fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Resource { limit, .. } => Some(limit),
            _ => None,
        }
    }

    /// Return the duplicate singular field name, if present.
    #[must_use]
    pub fn duplicate_singular_field(&self) -> Option<&str> {
        match &self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field.as_str()),
            _ => None,
        }
    }

    /// Return the missing required field, if present.
    #[must_use]
    pub fn missing_required_field(&self) -> Option<&str> {
        match &self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field.as_str()),
            _ => None,
        }
    }

    /// Return the canonicality failure reason, if present.
    #[must_use]
    pub fn noncanonical_reason(&self) -> Option<&str> {
        match &self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason.as_str()),
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
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::InvalidFormat(message) => formatter.write_str(message),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "theme required field {field} is absent")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "theme non-canonical representation: {reason}")
            },
            DecodeErrorKind::Resource {
                limit,
                observed,
                maximum,
            } => write!(
                formatter,
                "theme {limit} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "theme field {field} occurs more than once")
            },
            DecodeErrorKind::Projection => {
                formatter.write_str("theme Buffa projection disagreed with strict wire pass")
            },
            DecodeErrorKind::Allocation { amount } => {
                write!(formatter, "theme allocation of {amount} bytes failed")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self::invalid(format!("Buffa theme projection: {error}"))
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
}

impl DecodeOptions {
    /// Build an explicit finite input, field, work, and nesting policy.
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
        }
    }

    /// Construct a finite profile derived from one known source length.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(8).clamp(1, MAX_AUTOMATIC_LIMIT),
            bytes.saturating_mul(32).clamp(1, MAX_AUTOMATIC_LIMIT),
            8,
        )
        .with_max_output_bytes(bytes.saturating_mul(2).clamp(1, MAX_AUTOMATIC_LIMIT))
        .with_max_retained_bytes(bytes.saturating_mul(4).clamp(1, MAX_AUTOMATIC_LIMIT))
        .with_max_scratch_bytes(bytes.saturating_mul(2).clamp(1, MAX_AUTOMATIC_LIMIT))
    }

    /// Replace the source-byte ceiling.
    #[must_use]
    pub const fn with_max_message_bytes(mut self, maximum: usize) -> Self {
        self.max_message_bytes = maximum;
        self
    }

    /// Compatibility spelling for callers using input terminology.
    #[must_use]
    pub const fn with_max_input_bytes(self, maximum: usize) -> Self {
        self.with_max_message_bytes(maximum)
    }

    /// Replace the field-visit ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the aggregate work-byte ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Replace the candidate output-byte ceiling.
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

    /// Replace the maximum protobuf nesting depth.
    #[must_use]
    pub const fn with_max_depth(mut self, maximum: u32) -> Self {
        self.recursion_limit = maximum;
        self
    }

    /// Extend the aggregate field and work ceilings with a report that has
    /// already been charged by the same caller transaction.
    ///
    /// Source, output, retained, and scratch ceilings remain unchanged: the
    /// prior decode owns no candidate buffer. Callers must only pass a report
    /// produced for the same source and operation ledger.
    #[must_use]
    pub const fn with_additional_report(mut self, report: DecodeReport) -> Self {
        self.max_fields = self.max_fields.saturating_add(report.fields);
        self.max_work_bytes = self.max_work_bytes.saturating_add(report.work_bytes);
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Identity of the borrowed source region whose accounting was charged.
///
/// The pointer is safe to compare while a borrowed snapshot is alive: two
/// distinct live source regions cannot occupy the same address and length.
/// Keeping this token private prevents callers from manufacturing a report
/// for an unrelated source while preserving the report's `Copy` API.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct SourceToken {
    address: usize,
    length: usize,
}

impl SourceToken {
    fn for_source(source: &[u8]) -> Self {
        Self {
            address: source.as_ptr() as usize,
            length: source.len(),
        }
    }

    fn matches(self, source: &[u8]) -> bool {
        self == Self::for_source(source)
    }
}

/// Exact consumption of one successful theme decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    source_bytes: usize,
    source_token: SourceToken,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl DecodeReport {
    /// Source bytes inspected.
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    /// Alias for source_bytes.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.source_bytes
    }

    /// Strict field visits, including unknown fields and group delimiters.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate strict-plus-Buffa work charged by the decode.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum nested protobuf depth observed.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Logical allocations performed by the borrowed decode.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source bytes retained by the borrowed snapshot.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Scratch bytes allocated by the borrowed decode.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Borrowed semantic and source-preserving view of one complete theme.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ThemeArchiveSnapshot<'source> {
    source: &'source [u8],
    theme: &'source [u8],
    base: BaseSpans,
    identifier: Option<&'source str>,
}

/// One borrowed TSP reference selected from a theme base or preset family.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ThemeReferenceSnapshot {
    /// Native object identifier.
    pub identifier: u64,
    /// Optional deprecated type marker.
    pub deprecated_type: Option<i32>,
    /// Optional external-reference marker.
    pub deprecated_is_external: Option<bool>,
}

impl ThemeReferenceSnapshot {
    /// Construct a canonical reference with only its required identifier.
    #[must_use]
    pub const fn new(identifier: u64) -> Self {
        Self {
            identifier,
            deprecated_type: None,
            deprecated_is_external: None,
        }
    }

    /// Encode this reference as a bounded TSP.Reference payload.
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
        append_varint_field(&mut output, 1, self.identifier);
        if let Some(value) = self.deprecated_type {
            append_varint_field(&mut output, 2, value as i64 as u64);
        }
        if let Some(value) = self.deprecated_is_external {
            append_varint_field(&mut output, 3, u64::from(value));
        }
        Ok(output)
    }
}

/// Extension family whose repeated reference fields can be projected.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ThemeExtension {
    /// TSD drawing presets (field 100).
    Drawing,
    /// TSWP text presets (field 110).
    Text,
    /// TSCH chart presets (field 120).
    Chart,
    /// TST table presets (field 200).
    Table,
    /// TSA application presets (field 210).
    Application,
}

impl<'source> ThemeArchiveSnapshot<'source> {
    /// Return the complete wrapper bytes.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }

    /// Return the complete inner TSS.ThemeArchive bytes.
    #[must_use]
    pub const fn theme_raw(self) -> &'source [u8] {
        self.theme
    }

    /// Return the optional legacy stylesheet reference payload.
    #[must_use]
    pub fn legacy_stylesheet(self) -> Option<&'source [u8]> {
        self.base.legacy.map(|span| span.payload(self.theme))
    }

    /// Return the typed legacy stylesheet reference when present.
    #[must_use]
    pub fn legacy_stylesheet_reference(self) -> Option<ThemeReferenceSnapshot> {
        self.legacy_stylesheet().map(reference_snapshot)
    }

    /// Return the optional theme identifier.
    #[must_use]
    pub fn theme_identifier(self) -> Option<&'source str> {
        self.identifier
    }

    /// Return the optional document stylesheet reference payload.
    #[must_use]
    pub fn document_stylesheet(self) -> Option<&'source [u8]> {
        self.base.document.map(|span| span.payload(self.theme))
    }

    /// Return the typed document stylesheet reference when present.
    #[must_use]
    pub fn document_stylesheet_reference(self) -> Option<ThemeReferenceSnapshot> {
        self.document_stylesheet().map(reference_snapshot)
    }

    /// Return repeated old preset-replacement UUID payloads in source order.
    ///
    /// The bounded owner validates each payload as a complete TSP.UUID and
    /// preserves it exactly. New-package construction remains with the
    /// focused format builder that supplies these records.
    #[must_use]
    pub fn old_uuid_payloads(self) -> LengthDelimitedFields<'source> {
        LengthDelimitedFields::new(self.theme, OLD_PRESET_UUIDS_FIELD)
    }

    /// Return repeated new preset-replacement UUID payloads in source order.
    #[must_use]
    pub fn new_uuid_payloads(self) -> LengthDelimitedFields<'source> {
        LengthDelimitedFields::new(self.theme, NEW_PRESET_UUIDS_FIELD)
    }

    /// Return repeated color preset payloads in source order. Each payload is
    /// validated as a TSP.Color and remains source-owned for exact roundtrip.
    #[must_use]
    pub fn color_preset_payloads(self) -> LengthDelimitedFields<'source> {
        LengthDelimitedFields::new(self.theme, COLOR_PRESETS_FIELD)
    }

    /// Return the raw drawing preset extension payload.
    #[must_use]
    pub fn drawing_presets(self) -> Option<&'source [u8]> {
        self.extension(DRAWING_PRESETS_FIELD)
    }

    /// Return the raw text preset extension payload.
    #[must_use]
    pub fn text_presets(self) -> Option<&'source [u8]> {
        self.extension(TEXT_PRESETS_FIELD)
    }

    /// Return the raw chart preset extension payload.
    #[must_use]
    pub fn chart_presets(self) -> Option<&'source [u8]> {
        self.extension(CHART_PRESETS_FIELD)
    }

    /// Return the raw table preset extension payload.
    #[must_use]
    pub fn table_presets(self) -> Option<&'source [u8]> {
        self.extension(TABLE_PRESETS_FIELD)
    }

    /// Return the raw application preset extension payload.
    #[must_use]
    pub fn application_presets(self) -> Option<&'source [u8]> {
        self.extension(APPLICATION_PRESETS_FIELD)
    }

    /// Iterate one extension family's repeated TSP reference field.
    ///
    /// The selected field numbers are the native preset-list fields: drawing
    /// 4 through 9, text 1 through 8, chart 1, table 1 or 2, and application
    /// 1 or 2. Non-reference fields yield an empty iterator.
    #[must_use]
    pub fn extension_references(
        self,
        extension: ThemeExtension,
        field_number: u32,
    ) -> ReferenceFields<'source> {
        let payload = match extension {
            ThemeExtension::Drawing => self.drawing_presets(),
            ThemeExtension::Text => self.text_presets(),
            ThemeExtension::Chart => self.chart_presets(),
            ThemeExtension::Table => self.table_presets(),
            ThemeExtension::Application => self.application_presets(),
        };
        ReferenceFields::new(payload.unwrap_or_default(), field_number)
    }

    /// Iterate exact encoded wrapper fields that are not the required theme
    /// envelope.
    #[must_use]
    pub fn unknown_wrapper_fields(self) -> RawFields<'source> {
        RawFields::new(self.source, RawFieldScope::Wrapper)
    }

    /// Iterate exact encoded inner fields that are not an owned base field or
    /// one of the five known extension families.
    #[must_use]
    pub fn unknown_base_fields(self) -> RawFields<'source> {
        RawFields::new(self.theme, RawFieldScope::Base)
    }

    fn extension(self, field_number: u32) -> Option<&'source [u8]> {
        let index = EXTENSION_FIELDS
            .iter()
            .position(|number| *number == field_number)
            .expect("theme extension field");
        self.base.extensions[index].map(|span| span.payload(self.theme))
    }
}

/// A requested source-preserving theme mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ThemeArchiveWrite<'source> {
    legacy: FieldPatch<'source>,
    identifier: FieldPatch<'source>,
    document: FieldPatch<'source>,
    extensions: [FieldPatch<'source>; 5],
}

impl<'source> ThemeArchiveWrite<'source> {
    /// Build a no-op write that preserves every source span exactly.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            legacy: FieldPatch::Preserve,
            identifier: FieldPatch::Preserve,
            document: FieldPatch::Preserve,
            extensions: [FieldPatch::Preserve; 5],
        }
    }

    /// Replace the legacy stylesheet reference payload.
    #[must_use]
    pub const fn with_legacy_stylesheet(mut self, payload: &'source [u8]) -> Self {
        self.legacy = FieldPatch::Replace(payload);
        self
    }

    /// Clear the legacy stylesheet reference.
    #[must_use]
    pub const fn clear_legacy_stylesheet(mut self) -> Self {
        self.legacy = FieldPatch::Clear;
        self
    }

    /// Replace the UTF-8 theme identifier.
    #[must_use]
    pub fn with_theme_identifier(mut self, identifier: &'source str) -> Self {
        self.identifier = FieldPatch::Replace(identifier.as_bytes());
        self
    }

    /// Clear the theme identifier.
    #[must_use]
    pub const fn clear_theme_identifier(mut self) -> Self {
        self.identifier = FieldPatch::Clear;
        self
    }

    /// Replace the document stylesheet reference payload.
    #[must_use]
    pub const fn with_document_stylesheet(mut self, payload: &'source [u8]) -> Self {
        self.document = FieldPatch::Replace(payload);
        self
    }

    /// Clear the document stylesheet reference.
    #[must_use]
    pub const fn clear_document_stylesheet(mut self) -> Self {
        self.document = FieldPatch::Clear;
        self
    }

    /// Replace the TSD drawing preset extension payload.
    #[must_use]
    pub const fn with_drawing_presets(mut self, payload: &'source [u8]) -> Self {
        self.extensions[0] = FieldPatch::Replace(payload);
        self
    }

    /// Clear the TSD drawing preset extension.
    #[must_use]
    pub const fn clear_drawing_presets(mut self) -> Self {
        self.extensions[0] = FieldPatch::Clear;
        self
    }

    /// Replace the TSWP text preset extension payload.
    #[must_use]
    pub const fn with_text_presets(mut self, payload: &'source [u8]) -> Self {
        self.extensions[1] = FieldPatch::Replace(payload);
        self
    }

    /// Clear the TSWP text preset extension.
    #[must_use]
    pub const fn clear_text_presets(mut self) -> Self {
        self.extensions[1] = FieldPatch::Clear;
        self
    }

    /// Replace the TSCH chart preset extension payload.
    #[must_use]
    pub const fn with_chart_presets(mut self, payload: &'source [u8]) -> Self {
        self.extensions[2] = FieldPatch::Replace(payload);
        self
    }

    /// Clear the TSCH chart preset extension.
    #[must_use]
    pub const fn clear_chart_presets(mut self) -> Self {
        self.extensions[2] = FieldPatch::Clear;
        self
    }

    /// Replace the TST table preset extension payload.
    #[must_use]
    pub const fn with_table_presets(mut self, payload: &'source [u8]) -> Self {
        self.extensions[3] = FieldPatch::Replace(payload);
        self
    }

    /// Clear the TST table preset extension.
    #[must_use]
    pub const fn clear_table_presets(mut self) -> Self {
        self.extensions[3] = FieldPatch::Clear;
        self
    }

    /// Replace the TSA application preset extension payload.
    #[must_use]
    pub const fn with_application_presets(mut self, payload: &'source [u8]) -> Self {
        self.extensions[4] = FieldPatch::Replace(payload);
        self
    }

    /// Clear the TSA application preset extension.
    #[must_use]
    pub const fn clear_application_presets(mut self) -> Self {
        self.extensions[4] = FieldPatch::Clear;
        self
    }

    /// Rewrite one repeated TSP.Reference list in an existing extension.
    ///
    /// The rewrite remains part of the complete archive transaction. The
    /// source extension is scanned and the selected list is emitted directly
    /// into the final archive candidate, so no detached replacement buffer is
    /// needed between the extension and wrapper passes.
    #[must_use]
    pub fn with_reference_list(
        mut self,
        extension: ThemeExtension,
        field_number: u32,
        references: &'source [ThemeReferenceSnapshot],
    ) -> Self {
        self.extensions[extension_patch_index(extension)] = FieldPatch::ReferenceList {
            extension,
            field_number,
            references,
        };
        self
    }
}

impl Default for ThemeArchiveWrite<'_> {
    fn default() -> Self {
        Self::new()
    }
}

/// Bounded constructor for a fresh `TSS.ThemeArchive` wrapper.
///
/// Setters accept complete encoded payloads so a focused format package can
/// own the schema-specific preset records while this neutral owner remains
/// responsible for the envelope, repeated UUID/color records, limits, and
/// unknown-field preservation.  `encode` validates every supplied payload
/// before allocating the one candidate buffer.
#[derive(Debug)]
pub struct ThemeArchiveBuilder<'source> {
    legacy: Option<&'source [u8]>,
    identifier: Option<&'source str>,
    document: Option<&'source [u8]>,
    old_uuids: Option<&'source [&'source [u8]]>,
    new_uuids: Option<&'source [&'source [u8]]>,
    colors: Option<&'source [&'source [u8]]>,
    extensions: [Option<&'source [u8]>; 5],
    unknown_wrapper: Option<&'source [&'source [u8]]>,
    unknown_base: Option<&'source [&'source [u8]]>,
}

impl<'source> ThemeArchiveBuilder<'source> {
    /// Build an empty theme archive.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            legacy: None,
            identifier: None,
            document: None,
            old_uuids: None,
            new_uuids: None,
            colors: None,
            extensions: [None; 5],
            unknown_wrapper: None,
            unknown_base: None,
        }
    }

    /// Set the encoded legacy stylesheet `TSP.Reference` payload.
    #[must_use]
    pub const fn with_legacy_stylesheet(mut self, payload: &'source [u8]) -> Self {
        self.legacy = Some(payload);
        self
    }

    /// Set the UTF-8 theme identifier.
    #[must_use]
    pub fn with_theme_identifier(mut self, identifier: &'source str) -> Self {
        self.identifier = Some(identifier);
        self
    }

    /// Set the encoded document stylesheet `TSP.Reference` payload.
    #[must_use]
    pub const fn with_document_stylesheet(mut self, payload: &'source [u8]) -> Self {
        self.document = Some(payload);
        self
    }

    /// Set the borrowed encoded `TSP.UUID` old replacement list.
    #[must_use]
    pub const fn with_old_uuid_payloads(mut self, payloads: &'source [&'source [u8]]) -> Self {
        self.old_uuids = Some(payloads);
        self
    }

    /// Set the borrowed encoded `TSP.UUID` new replacement list.
    #[must_use]
    pub const fn with_new_uuid_payloads(mut self, payloads: &'source [&'source [u8]]) -> Self {
        self.new_uuids = Some(payloads);
        self
    }

    /// Set the borrowed encoded `TSP.Color` preset list.
    #[must_use]
    pub const fn with_color_preset_payloads(mut self, payloads: &'source [&'source [u8]]) -> Self {
        self.colors = Some(payloads);
        self
    }

    /// Set the complete TSD drawing preset extension payload.
    #[must_use]
    pub const fn with_drawing_presets(mut self, payload: &'source [u8]) -> Self {
        self.extensions[0] = Some(payload);
        self
    }

    /// Set the complete TSWP text preset extension payload.
    #[must_use]
    pub const fn with_text_presets(mut self, payload: &'source [u8]) -> Self {
        self.extensions[1] = Some(payload);
        self
    }

    /// Set the complete TSCH chart preset extension payload.
    #[must_use]
    pub const fn with_chart_presets(mut self, payload: &'source [u8]) -> Self {
        self.extensions[2] = Some(payload);
        self
    }

    /// Set the complete TST table preset extension payload.
    #[must_use]
    pub const fn with_table_presets(mut self, payload: &'source [u8]) -> Self {
        self.extensions[3] = Some(payload);
        self
    }

    /// Set the complete TSA application preset extension payload.
    #[must_use]
    pub const fn with_application_presets(mut self, payload: &'source [u8]) -> Self {
        self.extensions[4] = Some(payload);
        self
    }

    /// Set complete unknown wrapper fields for exact preservation.
    #[must_use]
    pub const fn with_unknown_wrapper_fields(mut self, fields: &'source [&'source [u8]]) -> Self {
        self.unknown_wrapper = Some(fields);
        self
    }

    /// Set complete unknown base fields for exact preservation.
    #[must_use]
    pub const fn with_unknown_base_fields(mut self, fields: &'source [&'source [u8]]) -> Self {
        self.unknown_base = Some(fields);
        self
    }

    /// Validate and encode one fresh wrapper under finite aggregate limits.
    ///
    /// The supplied records remain borrowed for the duration of this call;
    /// the returned bytes own the complete wrapper and all fields.  Exactly
    /// one output allocation is admitted and performed after all nested
    /// validation and size checks have passed.
    pub fn encode(self, options: DecodeOptions) -> Result<Vec<u8>, DecodeError> {
        validate_options(&[], options)?;
        if options.max_allocations < 1 {
            return Err(DecodeError::resource(
                DecodeLimit::Allocations,
                1,
                options.max_allocations,
            ));
        }
        let mut budget = Budget::new(&[], options);
        let mut inner_bytes = 0usize;

        if let Some(payload) = self.legacy {
            builder_payload(
                LEGACY_STYLESHEET_FIELD,
                payload,
                3,
                &mut budget,
                validate_reference,
            )?;
            add_builder_bytes(
                &mut inner_bytes,
                encoded_field_len(1, payload.len()),
                options,
            )?;
        }
        if let Some(identifier) = self.identifier {
            builder_simple_field(THEME_IDENTIFIER_FIELD, identifier.as_bytes(), &mut budget)?;
            add_builder_bytes(
                &mut inner_bytes,
                encoded_field_len(THEME_IDENTIFIER_FIELD, identifier.len()),
                options,
            )?;
        }
        if let Some(payload) = self.document {
            builder_payload(
                DOCUMENT_STYLESHEET_FIELD,
                payload,
                3,
                &mut budget,
                validate_reference,
            )?;
            add_builder_bytes(
                &mut inner_bytes,
                encoded_field_len(DOCUMENT_STYLESHEET_FIELD, payload.len()),
                options,
            )?;
        }
        for payload in self.old_uuids.unwrap_or_default().iter().copied() {
            builder_payload(
                OLD_PRESET_UUIDS_FIELD,
                payload,
                3,
                &mut budget,
                validate_uuid,
            )?;
            add_builder_bytes(
                &mut inner_bytes,
                encoded_field_len(OLD_PRESET_UUIDS_FIELD, payload.len()),
                options,
            )?;
        }
        for payload in self.new_uuids.unwrap_or_default().iter().copied() {
            builder_payload(
                NEW_PRESET_UUIDS_FIELD,
                payload,
                3,
                &mut budget,
                validate_uuid,
            )?;
            add_builder_bytes(
                &mut inner_bytes,
                encoded_field_len(NEW_PRESET_UUIDS_FIELD, payload.len()),
                options,
            )?;
        }
        for payload in self.colors.unwrap_or_default().iter().copied() {
            builder_payload(COLOR_PRESETS_FIELD, payload, 3, &mut budget, validate_color)?;
            add_builder_bytes(
                &mut inner_bytes,
                encoded_field_len(COLOR_PRESETS_FIELD, payload.len()),
                options,
            )?;
        }
        for (index, payload) in self.extensions.iter().copied().enumerate() {
            if let Some(payload) = payload {
                let number = EXTENSION_FIELDS[index];
                builder_payload(number, payload, 3, &mut budget, |payload, depth, budget| {
                    validate_theme_extension(payload, number, depth, budget)
                })?;
                add_builder_bytes(
                    &mut inner_bytes,
                    encoded_field_len(number, payload.len()),
                    options,
                )?;
            }
        }
        for field in self.unknown_base.unwrap_or_default().iter().copied() {
            validate_unknown_builder_field(field, RawFieldScope::Base, 2, &mut budget, options)?;
            add_builder_bytes(&mut inner_bytes, field.len(), options)?;
        }

        let mut output_bytes = encoded_field_len(THEME_SUPER_FIELD, inner_bytes);
        budget.field()?;
        for field in self.unknown_wrapper.unwrap_or_default().iter().copied() {
            validate_unknown_builder_field(field, RawFieldScope::Wrapper, 1, &mut budget, options)?;
            output_bytes = output_bytes
                .checked_add(field.len())
                .ok_or_else(|| output_overflow(options.max_output_bytes))?;
        }
        budget.charge_work(output_bytes)?;
        if output_bytes > options.max_output_bytes {
            return Err(DecodeError::resource(
                DecodeLimit::OutputBytes,
                output_bytes,
                options.max_output_bytes,
            ));
        }
        if output_bytes > options.max_retained_bytes {
            return Err(DecodeError::resource(
                DecodeLimit::RetainedBytes,
                output_bytes,
                options.max_retained_bytes,
            ));
        }
        if output_bytes > options.max_scratch_bytes {
            return Err(DecodeError::resource(
                DecodeLimit::ScratchBytes,
                output_bytes,
                options.max_scratch_bytes,
            ));
        }

        let mut output = Vec::new();
        output
            .try_reserve_exact(output_bytes)
            .map_err(|_| DecodeError::allocation(output_bytes))?;
        push_varint((u64::from(THEME_SUPER_FIELD) << 3) | 2, &mut output);
        push_varint(inner_bytes as u64, &mut output);
        if let Some(payload) = self.legacy {
            append_raw_field(&mut output, LEGACY_STYLESHEET_FIELD, payload);
        }
        if let Some(identifier) = self.identifier {
            append_raw_field(&mut output, THEME_IDENTIFIER_FIELD, identifier.as_bytes());
        }
        if let Some(payload) = self.document {
            append_raw_field(&mut output, DOCUMENT_STYLESHEET_FIELD, payload);
        }
        for payload in self.old_uuids.unwrap_or_default().iter().copied() {
            append_raw_field(&mut output, OLD_PRESET_UUIDS_FIELD, payload);
        }
        for payload in self.new_uuids.unwrap_or_default().iter().copied() {
            append_raw_field(&mut output, NEW_PRESET_UUIDS_FIELD, payload);
        }
        for payload in self.colors.unwrap_or_default().iter().copied() {
            append_raw_field(&mut output, COLOR_PRESETS_FIELD, payload);
        }
        for (index, payload) in self.extensions.iter().copied().enumerate() {
            if let Some(payload) = payload {
                append_raw_field(&mut output, EXTENSION_FIELDS[index], payload);
            }
        }
        for field in self.unknown_base.unwrap_or_default().iter().copied() {
            output.extend_from_slice(field);
        }
        for field in self.unknown_wrapper.unwrap_or_default().iter().copied() {
            output.extend_from_slice(field);
        }
        debug_assert_eq!(output.len(), output_bytes);
        Ok(output)
    }
}

fn builder_payload<F>(
    field_number: u32,
    payload: &[u8],
    depth: u32,
    budget: &mut Budget,
    validate: F,
) -> Result<(), DecodeError>
where
    F: FnOnce(&[u8], u32, &mut Budget) -> Result<(), DecodeError>,
{
    if payload.len() > budget.options.max_output_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::OutputBytes,
            payload.len(),
            budget.options.max_output_bytes,
        ));
    }
    budget.field()?;
    budget.charge_work(encoded_field_len(field_number, payload.len()))?;
    validate(payload, depth, budget)
}

fn builder_simple_field(
    field_number: u32,
    payload: &[u8],
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    if payload.len() > budget.options.max_output_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::OutputBytes,
            payload.len(),
            budget.options.max_output_bytes,
        ));
    }
    budget.field()?;
    budget.charge_work(encoded_field_len(field_number, payload.len()))
}

fn add_builder_bytes(
    total: &mut usize,
    amount: usize,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| output_overflow(options.max_output_bytes))?;
    if *total > options.max_output_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::OutputBytes,
            *total,
            options.max_output_bytes,
        ));
    }
    Ok(())
}

fn validate_unknown_builder_field(
    source: &[u8],
    scope: RawFieldScope,
    depth: u32,
    budget: &mut Budget,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if source.is_empty() {
        return Err(DecodeError::invalid("theme builder unknown field is empty"));
    }
    if source.len() > options.max_output_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::OutputBytes,
            source.len(),
            options.max_output_bytes,
        ));
    }
    // Read only the first tag before the bounded scan.  Calling `next_field`
    // here would recursively skip an unknown group without charging depth or
    // work; the scan below is the sole structural traversal.
    budget.charge_work(source.len().min(10))?;
    let (tag, _) = read_varint(source, 0)?;
    let number = u32::try_from(tag >> 3)
        .map_err(|_| DecodeError::invalid("theme builder unknown field overflows u32"))?;
    let known = match scope {
        RawFieldScope::Wrapper => number == THEME_SUPER_FIELD,
        RawFieldScope::Base => is_known_base_or_extension(number),
    };
    if known {
        return Err(DecodeError::invalid(
            "theme builder unknown field uses a known field number",
        ));
    }
    let mut top_fields = 0usize;
    let consumed = scan_fields(source, depth, None, budget, |_| {
        top_fields = top_fields
            .checked_add(1)
            .ok_or_else(|| field_overflow(options.max_fields))?;
        Ok(())
    })?;
    if consumed != source.len() || top_fields != 1 {
        return Err(DecodeError::invalid(
            "theme builder unknown field must contain exactly one field",
        ));
    }
    Ok(())
}

impl Default for ThemeArchiveBuilder<'_> {
    fn default() -> Self {
        Self::new()
    }
}

/// Exact finite requirements for one prepared theme rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    /// Candidate output bytes.
    pub output_bytes: usize,
    /// Aggregate field visits across preflight, emission, and readback.
    pub fields: usize,
    /// Aggregate strict-plus-Buffa work bytes.
    pub work_bytes: usize,
    /// Maximum nesting depth needed by the transaction.
    pub max_depth: u32,
    /// Logical output allocations.
    pub allocations: usize,
    /// Source and candidate bytes retained together.
    pub retained_bytes: usize,
    /// Candidate staging bytes.
    pub scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
    /// Convert requirements to exact execution ceilings.
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

/// Caller-provided ceilings checked before candidate allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    /// Candidate output bytes.
    pub output_bytes: usize,
    /// Aggregate field visits.
    pub fields: usize,
    /// Aggregate work bytes.
    pub work_bytes: usize,
    /// Maximum nesting depth.
    pub max_depth: u32,
    /// Logical output allocations.
    pub allocations: usize,
    /// Source and candidate bytes retained together.
    pub retained_bytes: usize,
    /// Candidate staging bytes.
    pub scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    /// Build exact limits from prepared requirements.
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

    /// Replace the maximum depth.
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

/// Exact accounting for one successful rewrite.
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

    /// Candidate bytes produced.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Aggregate field visits.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate work bytes.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum nested depth.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Logical output allocations.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source and candidate bytes retained together.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Candidate staging bytes.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Whether a selected source span changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
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

/// Prepared source-preserving rewrite. Execution checks caller limits before
/// allocating the candidate buffer.
#[derive(Debug)]
pub struct PreparedThemeArchiveRewrite<'source, 'write> {
    source: &'source [u8],
    write: ThemeArchiveWrite<'write>,
    preflight: Preflight<'source>,
    requirements: RewriteExecutionRequirements,
    changed: bool,
}

impl<'source, 'write> PreparedThemeArchiveRewrite<'source, 'write> {
    /// Return exact requirements computed before output allocation.
    #[must_use]
    pub const fn execution_requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute the prepared rewrite under caller-provided finite limits.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        ensure_requirements(self.requirements, limits)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| DecodeError::allocation(self.requirements.output_bytes))?;
        emit_wrapper(self.source, self.preflight, self.write, &mut output)?;
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::invalid(
                "theme rewrite size precharge disagreed with emission",
            ));
        }

        // Candidate readback is part of the same transaction.  Reusing the
        // automatic profile here would silently give the verification pass a
        // fresh field/work allowance unrelated to the requirements admitted
        // above.  Keep the exact aggregate ceilings for every pass.
        let readback_options = DecodeOptions::new(
            output.len().max(1),
            self.requirements.fields,
            self.requirements.work_bytes,
            self.requirements.max_depth.max(1),
        )
        .with_max_output_bytes(self.requirements.output_bytes)
        .with_max_allocations(self.requirements.allocations)
        .with_max_retained_bytes(self.requirements.retained_bytes)
        .with_max_scratch_bytes(self.requirements.scratch_bytes);
        let readback = decode_theme_archive_internal(&output, readback_options)?;
        let mut readback_budget = Budget::from_report(readback_options, readback.report);
        if !account_matches_write(readback.snapshot, self.write, &mut readback_budget)? {
            return Err(DecodeError::projection());
        }

        Ok(RewriteOutput {
            bytes: output,
            report: RewriteReport {
                input_bytes: self.source.len(),
                output_bytes: self.requirements.output_bytes,
                fields: self.requirements.fields,
                work_bytes: self.requirements.work_bytes,
                max_depth: self.requirements.max_depth,
                allocations: self.requirements.allocations,
                retained_bytes: self.requirements.retained_bytes,
                scratch_bytes: self.requirements.scratch_bytes,
                changed: self.changed,
            },
        })
    }
}

/// Decode a complete theme wrapper into a borrowed source-preserving view.
pub fn decode_theme_archive(
    source: &[u8],
    options: DecodeOptions,
) -> Result<ThemeArchiveSnapshot<'_>, DecodeError> {
    Ok(decode_theme_archive_internal(source, options)?.snapshot)
}

/// Decode a complete theme wrapper and return exact source accounting.
pub fn decode_theme_archive_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(ThemeArchiveSnapshot<'_>, DecodeReport), DecodeError> {
    let decoded = decode_theme_archive_internal(source, options)?;
    Ok((decoded.snapshot, decoded.report))
}

/// Admit an owned copy after a source decode that belongs to the same
/// transaction.
///
/// The borrowed snapshot and the returned copy coexist until the caller drops
/// the snapshot. Charge the copy against the source report before reserving
/// its buffer so work, retained bytes, scratch bytes, and the allocation event
/// share one finite ledger.
pub fn admit_theme_archive_copy(
    source: &[u8],
    report: DecodeReport,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    validate_options(source, options)?;
    if !report.source_token.matches(source) {
        return Err(DecodeError::invalid(
            "theme copy admission report does not belong to source",
        ));
    }
    let mut budget = Budget::new(source, options);
    budget.add_report(report)?;
    budget.charge_work(source.len())?;

    let retained_bytes = source
        .len()
        .checked_add(source.len())
        .ok_or_else(|| retained_overflow(options.max_retained_bytes))?;
    if retained_bytes > options.max_retained_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::RetainedBytes,
            retained_bytes,
            options.max_retained_bytes,
        ));
    }
    if source.len() > options.max_output_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::OutputBytes,
            source.len(),
            options.max_output_bytes,
        ));
    }
    if source.len() > options.max_scratch_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::ScratchBytes,
            source.len(),
            options.max_scratch_bytes,
        ));
    }
    if source.is_empty() {
        return Ok(Vec::new());
    }
    if options.max_allocations < 1 {
        return Err(DecodeError::resource(
            DecodeLimit::Allocations,
            1,
            options.max_allocations,
        ));
    }

    let mut copy = Vec::new();
    copy.try_reserve_exact(source.len())
        .map_err(|_| DecodeError::allocation(source.len()))?;
    copy.extend_from_slice(source);
    Ok(copy)
}

/// Project several native repeated-reference lists under one decode and
/// projection ledger.
///
/// The selector slice and route projection are borrowed from the caller, so
/// selector staging does not allocate before the codec has admitted the
/// source traversal and result vector.
pub fn project_theme_references<T, F>(
    source: &[u8],
    selectors: &[T],
    route: F,
    options: DecodeOptions,
) -> Result<Vec<ThemeReferenceSnapshot>, DecodeError>
where
    T: Copy,
    F: Fn(T) -> (ThemeExtension, u32),
{
    for selector in selectors.iter().copied() {
        let (extension, field_number) = route(selector);
        if !extension_reference_field(extension, field_number) {
            return Err(DecodeError::invalid(format!(
                "field {field_number} is not a TSP.Reference list in this theme extension"
            )));
        }
    }
    let decoded = decode_theme_archive_internal(source, options)?;
    let mut budget = Budget::from_report(options, decoded.report);

    let count = account_reference_projection(
        decoded.snapshot,
        selectors.iter().copied(),
        &route,
        &mut budget,
    )?;
    let repeated_count = account_reference_projection(
        decoded.snapshot,
        selectors.iter().copied(),
        &route,
        &mut budget,
    )?;
    if repeated_count != count {
        return Err(DecodeError::projection());
    }
    // The result loop below reparses each selected extension list while it
    // materializes borrowed snapshots. Charge that replay before reserving
    // the owned result vector; the two cardinality passes above do not make
    // the later iterator free.
    let materialized_count = account_reference_projection(
        decoded.snapshot,
        selectors.iter().copied(),
        &route,
        &mut budget,
    )?;
    if materialized_count != count {
        return Err(DecodeError::projection());
    }
    let projection_bytes = count
        .checked_mul(size_of::<ThemeReferenceSnapshot>())
        .ok_or_else(|| work_overflow(options.max_work_bytes))?;
    budget.charge_work(projection_bytes)?;
    if projection_bytes > options.max_scratch_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::ScratchBytes,
            projection_bytes,
            options.max_scratch_bytes,
        ));
    }
    let retained_bytes = source
        .len()
        .checked_add(projection_bytes)
        .ok_or_else(|| retained_overflow(options.max_retained_bytes))?;
    if retained_bytes > options.max_retained_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::RetainedBytes,
            retained_bytes,
            options.max_retained_bytes,
        ));
    }
    if count == 0 {
        return Ok(Vec::new());
    }
    if options.max_allocations < 1 {
        return Err(DecodeError::resource(
            DecodeLimit::Allocations,
            1,
            options.max_allocations,
        ));
    }

    let mut references = Vec::new();
    references
        .try_reserve_exact(count)
        .map_err(|_| DecodeError::allocation(projection_bytes))?;
    for selector in selectors.iter().copied() {
        let (extension, field_number) = route(selector);
        references.extend(
            decoded
                .snapshot
                .extension_references(extension, field_number),
        );
    }
    if references.len() != count {
        return Err(DecodeError::projection());
    }
    Ok(references)
}

/// Return the first native reference without materializing the selected list.
pub fn first_theme_reference(
    source: &[u8],
    extension: ThemeExtension,
    field_number: u32,
    options: DecodeOptions,
) -> Result<Option<ThemeReferenceSnapshot>, DecodeError> {
    if !extension_reference_field(extension, field_number) {
        return Err(DecodeError::invalid(format!(
            "field {field_number} is not a TSP.Reference list in this theme extension"
        )));
    }
    let decoded = decode_theme_archive_internal(source, options)?;
    let mut budget = Budget::from_report(options, decoded.report);
    let Some(payload) = snapshot_extension_payload(decoded.snapshot, extension) else {
        return Ok(None);
    };
    let mut cursor = 0usize;
    while cursor < payload.len() {
        let field = next_field(payload, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme extension source changed after scan"))?;
        budget.field()?;
        budget.charge_work(field.span.end - field.span.start)?;
        if field.number == field_number {
            budget.charge_work(field.span.payload(payload).len())?;
            return Ok(Some(reference_snapshot(field.span.payload(payload))));
        }
    }
    Ok(None)
}

/// Compare one repeated native reference list while debiting the comparison
/// replay into an existing decode report. Owners that need an exact no-op
/// fast path can pass the returned report into a subsequent rewrite without
/// leaving the `ReferenceFields` walk outside the transaction ledger.
pub fn theme_reference_list_matches_with_report(
    snapshot: ThemeArchiveSnapshot<'_>,
    extension: ThemeExtension,
    field_number: u32,
    references: &[ThemeReferenceSnapshot],
    options: DecodeOptions,
    prior_report: DecodeReport,
) -> Result<(bool, DecodeReport), DecodeError> {
    if !extension_reference_field(extension, field_number) {
        return Err(DecodeError::invalid(format!(
            "field {field_number} is not a TSP.Reference list in this theme extension"
        )));
    }
    validate_options(snapshot.raw(), options)?;
    if !prior_report.source_token.matches(snapshot.raw()) {
        return Err(DecodeError::invalid(
            "theme reference comparison report does not belong to snapshot",
        ));
    }
    let mut budget = Budget::from_report(options, prior_report);
    let matches = account_reference_fields_match(
        snapshot_extension_payload(snapshot, extension),
        field_number,
        references,
        &mut budget,
    )?;
    Ok((matches, budget.decode_report()))
}

/// Prepare a complete source-preserving theme rewrite.
pub fn prepare_theme_archive_rewrite<'source, 'write>(
    source: &'source [u8],
    write: ThemeArchiveWrite<'write>,
    options: DecodeOptions,
) -> Result<PreparedThemeArchiveRewrite<'source, 'write>, DecodeError> {
    prepare_theme_archive_rewrite_with_prior_report(source, write, options, None)
}

/// Prepare a complete rewrite after a bounded source read that belongs to the
/// same caller transaction.
///
/// The prior report is debited into the same field/work ledger before the
/// rewrite is sized. This is used by owners that must inspect a selected list
/// for an exact no-op before admitting the candidate rewrite.
pub fn prepare_theme_archive_rewrite_with_prior_report<'source, 'write>(
    source: &'source [u8],
    write: ThemeArchiveWrite<'write>,
    options: DecodeOptions,
    prior_report: Option<DecodeReport>,
) -> Result<PreparedThemeArchiveRewrite<'source, 'write>, DecodeError> {
    let decoded = decode_theme_archive_internal(source, options)?;
    let mut budget = Budget::from_report(options, decoded.report);
    if let Some(prior_report) = prior_report {
        if !prior_report.source_token.matches(source) {
            return Err(DecodeError::invalid(
                "theme rewrite report does not belong to source",
            ));
        }
        budget.add_report(prior_report)?;
    }
    validate_write(write, &mut budget)?;
    let source_matches = account_matches_write(decoded.snapshot, write, &mut budget)?;

    let (inner_bytes, output_bytes) =
        measure_rewrite(source, decoded.preflight, write, &mut budget)?;
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::OutputBytes,
            output_bytes,
            options.max_output_bytes,
        ));
    }

    // The sizing pass above is followed by a different emission shape: the
    // inner message is measured once more, then the wrapper and inner message
    // are copied.  Finally the complete candidate is decoded again through
    // the strict pass and the Buffa projection.  Account for all of those
    // traversals before the output Vec exists.  The old implementation only
    // charged a loose output-byte multiplier and left these field visits to
    // execute/readback, where a tight caller ceiling could fail too late.
    account_emission(
        source,
        decoded.preflight.theme,
        write,
        output_bytes,
        &mut budget,
    )?;
    account_candidate_readback(source, decoded.preflight, write, inner_bytes, &mut budget)?;
    account_candidate_matches_write(decoded.preflight.theme, write, &mut budget)?;
    budget.buffa(output_bytes)?;
    budget.buffa(inner_bytes)?;

    let fields = budget.fields;
    let work_bytes = budget.work_bytes;
    if fields > options.max_fields {
        return Err(DecodeError::resource(
            DecodeLimit::Fields,
            fields,
            options.max_fields,
        ));
    }
    if work_bytes > options.max_work_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::WorkBytes,
            work_bytes,
            options.max_work_bytes,
        ));
    }
    let retained_bytes = source
        .len()
        .checked_add(output_bytes)
        .ok_or_else(|| retained_overflow(options.max_retained_bytes))?;
    if retained_bytes > options.max_retained_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::RetainedBytes,
            retained_bytes,
            options.max_retained_bytes,
        ));
    }
    if output_bytes > options.max_scratch_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::ScratchBytes,
            output_bytes,
            options.max_scratch_bytes,
        ));
    }
    if options.max_allocations < 1 {
        return Err(DecodeError::resource(
            DecodeLimit::Allocations,
            1,
            options.max_allocations,
        ));
    }

    let _ = inner_bytes;
    Ok(PreparedThemeArchiveRewrite {
        source,
        write,
        preflight: decoded.preflight,
        requirements: RewriteExecutionRequirements {
            output_bytes,
            fields,
            work_bytes,
            max_depth: budget.max_depth,
            allocations: 1,
            retained_bytes,
            scratch_bytes: output_bytes,
        },
        // A reference-list patch can be a semantic no-op.  In that case the
        // source-preserving emitter retains the original records (including
        // nested unknown fields), so the report must expose an exact no-op
        // rather than treating every selector patch as a change.
        changed: !source_matches,
    })
}

/// Rewrite a complete theme wrapper and return candidate bytes.
pub fn rewrite_theme_archive(
    source: &[u8],
    write: ThemeArchiveWrite<'_>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    Ok(rewrite_theme_archive_with_report(source, write, options)?.0)
}

/// Rewrite a complete theme wrapper and return candidate bytes plus exact
/// aggregate accounting.
pub fn rewrite_theme_archive_with_report(
    source: &[u8],
    write: ThemeArchiveWrite<'_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let prepared = prepare_theme_archive_rewrite(source, write, options)?;
    let requirements = prepared.execution_requirements();
    let output = prepared.execute(RewriteExecutionLimits::exact(requirements))?;
    Ok((output.bytes, output.report))
}

/// Rewrite a complete theme wrapper after a bounded source read that belongs
/// to the same caller transaction.
pub fn rewrite_theme_archive_with_prior_report(
    source: &[u8],
    write: ThemeArchiveWrite<'_>,
    options: DecodeOptions,
    prior_report: DecodeReport,
) -> Result<Vec<u8>, DecodeError> {
    let prepared = prepare_theme_archive_rewrite_with_prior_report(
        source,
        write,
        options,
        Some(prior_report),
    )?;
    let requirements = prepared.execution_requirements();
    Ok(prepared
        .execute(RewriteExecutionLimits::exact(requirements))?
        .into_bytes())
}

/// One source-preserving emission selected while rewriting a repeated
/// TSP.Reference field.
#[derive(Debug, Clone, Copy)]
enum ReferenceEmission<'source> {
    Raw(&'source [u8]),
    Canonical(ThemeReferenceSnapshot),
}

/// Find the source record that supplies one requested reference.
///
/// A reference list can be reordered while its individual `TSP.Reference`
/// records carry unknown nested fields. Matching by the requested semantic
/// value and its occurrence number lets the rewrite retain those complete
/// source records for arbitrary permutations and repeated values. The helper
/// deliberately borrows the source; it does not build an unbounded side table
/// before the rewrite budget has admitted the traversal.
fn raw_reference_for_desired<'source>(
    source: &'source [u8],
    field_number: u32,
    references: &[ThemeReferenceSnapshot],
    desired_index: usize,
    budget: &mut Option<&mut Budget>,
) -> Result<Option<&'source [u8]>, DecodeError> {
    let desired = references[desired_index];
    let mut desired_occurrence = 0usize;
    for reference in &references[..desired_index] {
        if let Some(budget) = budget.as_deref_mut() {
            // Matching the requested occurrence scans the already requested
            // prefix for every desired slot. This is intentionally charged:
            // duplicate-heavy reorder requests otherwise hide an O(n²)
            // comparison pass outside the aggregate work ledger.
            budget.charge_work(size_of::<ThemeReferenceSnapshot>())?;
        }
        if *reference == desired {
            desired_occurrence = desired_occurrence
                .checked_add(1)
                .ok_or_else(|| field_overflow(usize::MAX))?;
        }
    }
    let mut occurrence = 0usize;
    let mut cursor = 0usize;
    while cursor < source.len() {
        let field = next_field(source, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme extension source changed after scan"))?;
        if let Some(budget) = budget.as_deref_mut() {
            // This pass is a source traversal in addition to the emitting
            // pass. Charge it before any caller-owned output is allocated.
            budget.field()?;
            budget.charge_work(field.span.end - field.span.start)?;
        }
        if field.number != field_number {
            continue;
        }
        if reference_snapshot(field.span.payload(source)) != desired {
            continue;
        }
        if occurrence == desired_occurrence {
            return Ok(Some(&source[field.span.start..field.span.end]));
        }
        occurrence = occurrence
            .checked_add(1)
            .ok_or_else(|| field_overflow(usize::MAX))?;
    }
    Ok(None)
}

fn walk_reference_rewrite<'source, F>(
    source: &'source [u8],
    field_number: u32,
    references: &[ThemeReferenceSnapshot],
    mut budget: Option<&mut Budget>,
    mut emit: F,
) -> Result<(usize, usize), DecodeError>
where
    F: FnMut(ReferenceEmission<'source>) -> Result<(), DecodeError>,
{
    let mut source_fields = 0usize;
    let mut source_wire_bytes = 0usize;
    let mut desired = 0usize;
    let mut cursor = 0usize;
    while cursor < source.len() {
        let field = next_field(source, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme extension source changed after scan"))?;
        source_fields = source_fields
            .checked_add(1)
            .ok_or_else(|| field_overflow(usize::MAX))?;
        let span_bytes = field.span.end - field.span.start;
        source_wire_bytes = source_wire_bytes
            .checked_add(span_bytes)
            .ok_or_else(|| work_overflow(usize::MAX))?;
        if let Some(budget) = budget.as_deref_mut() {
            budget.field()?;
            budget.charge_work(span_bytes)?;
        }

        let raw = &source[field.span.start..field.span.end];
        if field.number == field_number {
            if desired < references.len() {
                // The requested sequence, rather than source order, controls
                // the output slot. If a semantically matching source record
                // exists, retain its complete raw bytes even when it came
                // from a different source slot.
                if let Some(raw) = raw_reference_for_desired(
                    source,
                    field_number,
                    references,
                    desired,
                    &mut budget,
                )? {
                    emit(ReferenceEmission::Raw(raw))?;
                } else {
                    emit(ReferenceEmission::Canonical(references[desired]))?;
                }
            }
            desired = desired
                .checked_add(1)
                .ok_or_else(|| field_overflow(usize::MAX))?;
        } else {
            emit(ReferenceEmission::Raw(raw))?;
        }
    }
    while desired < references.len() {
        if let Some(raw) =
            raw_reference_for_desired(source, field_number, references, desired, &mut budget)?
        {
            emit(ReferenceEmission::Raw(raw))?;
        } else {
            emit(ReferenceEmission::Canonical(references[desired]))?;
        }
        desired = desired
            .checked_add(1)
            .ok_or_else(|| field_overflow(usize::MAX))?;
    }
    Ok((source_fields, source_wire_bytes))
}

fn reference_emission_len(field_number: u32, emission: ReferenceEmission<'_>) -> usize {
    match emission {
        ReferenceEmission::Raw(bytes) => bytes.len(),
        ReferenceEmission::Canonical(reference) => {
            encoded_reference_field_len(field_number, reference)
        },
    }
}

fn measure_reference_list_payload(
    source: &[u8],
    field_number: u32,
    references: &[ThemeReferenceSnapshot],
    budget: Option<&mut Budget>,
) -> Result<usize, DecodeError> {
    let mut output_bytes = 0usize;
    let mut canonical_fields = 0usize;
    let mut budget = budget;
    walk_reference_rewrite(
        source,
        field_number,
        references,
        budget.as_deref_mut(),
        |emission| {
            if matches!(emission, ReferenceEmission::Canonical(_)) {
                canonical_fields = canonical_fields
                    .checked_add(1)
                    .ok_or_else(|| field_overflow(usize::MAX))?;
            }
            output_bytes = output_bytes
                .checked_add(reference_emission_len(field_number, emission))
                .ok_or_else(|| output_overflow(usize::MAX))?;
            Ok(())
        },
    )?;
    if let Some(budget) = budget {
        for _ in 0..canonical_fields {
            budget.field()?;
        }
        // The caller will later reserve/copy this exact payload. Account its
        // encoding work while the source is still borrowed and no output Vec
        // exists.
        budget.charge_work(output_bytes)?;
    }
    Ok(output_bytes)
}

fn emit_reference_list_payload(
    source: &[u8],
    field_number: u32,
    references: &[ThemeReferenceSnapshot],
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    walk_reference_rewrite(source, field_number, references, None, |emission| {
        match emission {
            ReferenceEmission::Raw(bytes) => output.extend_from_slice(bytes),
            ReferenceEmission::Canonical(reference) => {
                append_reference_field(output, field_number, reference);
            },
        }
        Ok(())
    })?;
    Ok(())
}

/// Replace one repeated TSP.Reference list inside a theme extension.
///
/// The input is the raw payload returned by one of the extension accessors.
/// Every other extension field, including unknown and complex fill/shadow
/// records, is copied byte-for-byte. Existing selected records are also
/// copied byte-for-byte when their known reference fields match the requested
/// source-order sequence; this preserves unknown nested fields in untouched
/// `TSP.Reference` records while still allowing insertions, removals, and
/// replacements.
pub fn rewrite_theme_extension_reference_list(
    source: &[u8],
    extension: ThemeExtension,
    field_number: u32,
    references: &[ThemeReferenceSnapshot],
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let extension_field = extension_field(extension);
    if !extension_reference_field(extension, field_number) {
        return Err(DecodeError::invalid(format!(
            "field {field_number} is not a TSP.Reference list in this theme extension"
        )));
    }
    validate_options(source, options)?;
    let mut budget = Budget::new(source, options);
    validate_theme_extension(source, extension_field, 1, &mut budget)?;
    if options.max_allocations < 1 {
        return Err(DecodeError::resource(
            DecodeLimit::Allocations,
            1,
            options.max_allocations,
        ));
    }

    // Admit the replacement values and the first source walk before sizing
    // the candidate. The second source walk is the fallible-free emitter.
    for reference in references {
        let encoded = encoded_reference_field_len(field_number, *reference);
        budget.field()?;
        budget.charge_work(encoded)?;
    }

    let mut output_bytes = 0usize;
    walk_reference_rewrite(
        source,
        field_number,
        references,
        Some(&mut budget),
        |emission| {
            let bytes = match emission {
                ReferenceEmission::Raw(bytes) => bytes.len(),
                ReferenceEmission::Canonical(reference) => {
                    encoded_reference_field_len(field_number, reference)
                },
            };
            output_bytes = output_bytes
                .checked_add(bytes)
                .ok_or_else(|| output_overflow(options.max_output_bytes))?;
            if output_bytes > options.max_output_bytes {
                return Err(DecodeError::resource(
                    DecodeLimit::OutputBytes,
                    output_bytes,
                    options.max_output_bytes,
                ));
            }
            Ok(())
        },
    )?;

    // The emitter reparses the source, repeats the duplicate/reorder matcher,
    // and walks the replacement slice once more. Reuse the budgeted sizing
    // helper so every source scan, occurrence-prefix comparison, canonical
    // field, and output byte is admitted before the output Vec exists.
    let replay_bytes =
        measure_reference_list_payload(source, field_number, references, Some(&mut budget))?;
    if replay_bytes != output_bytes {
        return Err(DecodeError::invalid(
            "theme extension rewrite replay disagreed with sizing",
        ));
    }
    // The final emitter still walks the source once more after the candidate
    // buffer is reserved. Replay that walk under the same ledger before the
    // reservation so the fallible-free copy cannot escape aggregate work
    // admission.
    let final_bytes =
        measure_reference_list_payload(source, field_number, references, Some(&mut budget))?;
    if final_bytes != output_bytes {
        return Err(DecodeError::invalid(
            "theme extension rewrite emission disagreed with sizing",
        ));
    }

    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::OutputBytes,
            output_bytes,
            options.max_output_bytes,
        ));
    }
    if output_bytes > options.max_scratch_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::ScratchBytes,
            output_bytes,
            options.max_scratch_bytes,
        ));
    }
    let maximum_retained = source
        .len()
        .checked_add(output_bytes)
        .ok_or_else(|| retained_overflow(options.max_retained_bytes))?;
    if maximum_retained > options.max_retained_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::RetainedBytes,
            maximum_retained,
            options.max_retained_bytes,
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_bytes)
        .map_err(|_| DecodeError::allocation(output_bytes))?;
    walk_reference_rewrite(source, field_number, references, None, |emission| {
        match emission {
            ReferenceEmission::Raw(bytes) => output.extend_from_slice(bytes),
            ReferenceEmission::Canonical(reference) => {
                append_reference_field(&mut output, field_number, reference);
            },
        }
        Ok(())
    })?;
    if output.len() != output_bytes {
        return Err(DecodeError::invalid(
            "theme extension rewrite size precharge disagreed with emission",
        ));
    }
    if output.len() > options.max_output_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::OutputBytes,
            output.len(),
            options.max_output_bytes,
        ));
    }
    Ok(output)
}

/// Iterator over exact encoded unknown fields.
#[derive(Debug, Clone, Copy)]
pub struct RawFields<'source> {
    source: &'source [u8],
    cursor: usize,
    scope: RawFieldScope,
}

impl<'source> RawFields<'source> {
    fn new(source: &'source [u8], scope: RawFieldScope) -> Self {
        Self {
            source,
            cursor: 0,
            scope,
        }
    }
}

impl<'source> Iterator for RawFields<'source> {
    type Item = &'source [u8];

    fn next(&mut self) -> Option<Self::Item> {
        while self.cursor < self.source.len() {
            let field = next_field(self.source, &mut self.cursor)?;
            let raw = &self.source[field.span.start..field.span.end];
            let known = match self.scope {
                RawFieldScope::Wrapper => field.number == THEME_SUPER_FIELD,
                RawFieldScope::Base => is_known_base_or_extension(field.number),
            };
            if !known {
                return Some(raw);
            }
        }
        None
    }
}

/// Iterator over repeated length-delimited payloads for one base field.
#[derive(Debug, Clone, Copy)]
pub struct LengthDelimitedFields<'source> {
    source: &'source [u8],
    cursor: usize,
    number: u32,
}

impl<'source> LengthDelimitedFields<'source> {
    fn new(source: &'source [u8], number: u32) -> Self {
        Self {
            source,
            cursor: 0,
            number,
        }
    }
}

impl<'source> Iterator for LengthDelimitedFields<'source> {
    type Item = &'source [u8];

    fn next(&mut self) -> Option<Self::Item> {
        while self.cursor < self.source.len() {
            let field = next_field(self.source, &mut self.cursor)?;
            if field.number == self.number && field.wire_type == 2 {
                return Some(&self.source[field.span.payload_start..field.span.payload_end]);
            }
        }
        None
    }
}

/// Iterator over typed TSP references in one selected extension field.
#[derive(Debug, Clone, Copy)]
pub struct ReferenceFields<'source> {
    source: &'source [u8],
    cursor: usize,
    number: u32,
}

impl<'source> ReferenceFields<'source> {
    fn new(source: &'source [u8], number: u32) -> Self {
        Self {
            source,
            cursor: 0,
            number,
        }
    }
}

impl Iterator for ReferenceFields<'_> {
    type Item = ThemeReferenceSnapshot;

    fn next(&mut self) -> Option<Self::Item> {
        while self.cursor < self.source.len() {
            let field = next_field(self.source, &mut self.cursor)?;
            if field.number == self.number && field.wire_type == 2 {
                return Some(reference_snapshot(
                    &self.source[field.span.payload_start..field.span.payload_end],
                ));
            }
        }
        None
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RawFieldScope {
    Wrapper,
    Base,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FieldPatch<'source> {
    Preserve,
    Clear,
    Replace(&'source [u8]),
    ReferenceList {
        extension: ThemeExtension,
        field_number: u32,
        references: &'source [ThemeReferenceSnapshot],
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Span {
    start: usize,
    payload_start: usize,
    payload_end: usize,
    end: usize,
}

impl Span {
    fn payload(self, source: &[u8]) -> &[u8] {
        &source[self.payload_start..self.payload_end]
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct WireField {
    number: u32,
    wire_type: u8,
    span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
struct BaseSpans {
    legacy: Option<Span>,
    identifier: Option<Span>,
    document: Option<Span>,
    extensions: [Option<Span>; 5],
}

#[derive(Debug, Clone, Copy)]
struct Preflight<'source> {
    theme: &'source [u8],
    base: BaseSpans,
}

#[derive(Debug, Clone, Copy)]
struct Decoded<'source> {
    snapshot: ThemeArchiveSnapshot<'source>,
    preflight: Preflight<'source>,
    report: DecodeReport,
}

fn decode_theme_archive_internal(
    source: &[u8],
    options: DecodeOptions,
) -> Result<Decoded<'_>, DecodeError> {
    validate_options(source, options)?;
    let mut budget = Budget::new(source, options);
    let preflight = preflight_theme(source, &mut budget)?;

    budget.buffa(source.len())?;
    let wrapper: projection::ThemeWrapperLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    if !wrapper.has_super() || wrapper.super_ != preflight.theme {
        return Err(DecodeError::projection());
    }

    budget.buffa(preflight.theme.len())?;
    let theme: projection::ThemeArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(preflight.theme)
        .map_err(DecodeError::from)?;
    if theme.legacy_stylesheet != payload(preflight.theme, preflight.base.legacy)
        || theme.document_stylesheet != payload(preflight.theme, preflight.base.document)
        || theme.theme_identifier
            != preflight.base.identifier.map(|span| {
                str::from_utf8(span.payload(preflight.theme)).expect("validated theme identifier")
            })
        || theme.drawing_presets != payload(preflight.theme, preflight.base.extensions[0])
        || theme.text_presets != payload(preflight.theme, preflight.base.extensions[1])
        || theme.chart_presets != payload(preflight.theme, preflight.base.extensions[2])
        || theme.table_presets != payload(preflight.theme, preflight.base.extensions[3])
        || theme.application_presets != payload(preflight.theme, preflight.base.extensions[4])
    {
        return Err(DecodeError::projection());
    }

    let report = budget.decode_report();
    let snapshot = ThemeArchiveSnapshot {
        source,
        theme: preflight.theme,
        base: preflight.base,
        identifier: preflight.base.identifier.map(|span| {
            str::from_utf8(span.payload(preflight.theme)).expect("validated theme identifier")
        }),
    };
    Ok(Decoded {
        snapshot,
        preflight,
        report,
    })
}

fn preflight_theme<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<Preflight<'source>, DecodeError> {
    let mut theme_span = None;
    scan_fields(source, 1, None, budget, |field| {
        if field.number == THEME_SUPER_FIELD {
            if field.wire_type != 2 {
                return Err(DecodeError::invalid(
                    "iWork theme wrapper field 1 is not length-delimited",
                ));
            }
            if theme_span.replace(field.span).is_some() {
                return Err(DecodeError::duplicate("TSS.ThemeArchive wrapper super"));
            }
        }
        Ok(())
    })?;
    let theme_span = theme_span.ok_or_else(|| {
        DecodeError::invalid("iWork theme wrapper must contain exactly one field 1")
    })?;
    let theme = theme_span.payload(source);
    let base = preflight_base(theme, budget)?;
    Ok(Preflight { theme, base })
}

fn preflight_base(theme: &[u8], budget: &mut Budget) -> Result<BaseSpans, DecodeError> {
    let mut base = BaseSpans::default();
    scan_fields(theme, 2, None, budget, |field| {
        match field.number {
            LEGACY_STYLESHEET_FIELD => {
                require_wire(field, 2, "TSS.ThemeArchive.legacy_stylesheet")?;
                if base.legacy.replace(field.span).is_some() {
                    return Err(DecodeError::duplicate("TSS.ThemeArchive.legacy_stylesheet"));
                }
            },
            THEME_IDENTIFIER_FIELD => {
                require_wire(field, 2, "TSS.ThemeArchive.theme_identifier")?;
                if base.identifier.replace(field.span).is_some() {
                    return Err(DecodeError::duplicate("TSS.ThemeArchive.theme_identifier"));
                }
                let payload = &theme[field.span.payload_start..field.span.payload_end];
                str::from_utf8(payload).map_err(|_| {
                    DecodeError::invalid("TSS.ThemeArchive.theme_identifier is not UTF-8")
                })?;
            },
            DOCUMENT_STYLESHEET_FIELD => {
                require_wire(field, 2, "TSS.ThemeArchive.document_stylesheet")?;
                if base.document.replace(field.span).is_some() {
                    return Err(DecodeError::duplicate(
                        "TSS.ThemeArchive.document_stylesheet",
                    ));
                }
            },
            OLD_PRESET_UUIDS_FIELD | NEW_PRESET_UUIDS_FIELD => {
                require_wire(field, 2, "TSS.ThemeArchive preset-replacement UUID")?;
            },
            COLOR_PRESETS_FIELD => {
                require_wire(field, 2, "TSS.ThemeArchive.color_presets")?;
            },
            DRAWING_PRESETS_FIELD
            | TEXT_PRESETS_FIELD
            | CHART_PRESETS_FIELD
            | TABLE_PRESETS_FIELD
            | APPLICATION_PRESETS_FIELD => {
                require_wire(field, 2, "TSS.ThemeArchive extension")?;
                let index = extension_index(field.number).expect("known extension");
                if base.extensions[index].replace(field.span).is_some() {
                    return Err(DecodeError::duplicate(format!(
                        "TSS.ThemeArchive extension field {}",
                        field.number
                    )));
                }
            },
            _ => {},
        }
        Ok(())
    })?;

    let mut cursor = 0usize;
    while cursor < theme.len() {
        let field = next_field(theme, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme source changed after preflight"))?;
        let payload = &theme[field.span.payload_start..field.span.payload_end];
        match field.number {
            LEGACY_STYLESHEET_FIELD => validate_reference(payload, 3, budget)?,
            DOCUMENT_STYLESHEET_FIELD => validate_reference(payload, 3, budget)?,
            OLD_PRESET_UUIDS_FIELD | NEW_PRESET_UUIDS_FIELD => validate_uuid(payload, 3, budget)?,
            COLOR_PRESETS_FIELD => validate_color(payload, 3, budget)?,
            DRAWING_PRESETS_FIELD
            | TEXT_PRESETS_FIELD
            | CHART_PRESETS_FIELD
            | TABLE_PRESETS_FIELD
            | APPLICATION_PRESETS_FIELD => {
                validate_theme_extension(payload, field.number, 3, budget)?;
            },
            _ => {},
        }
    }
    Ok(base)
}

fn validate_write(write: ThemeArchiveWrite<'_>, budget: &mut Budget) -> Result<(), DecodeError> {
    validate_patch_budget(write.legacy, budget)?;
    match write.legacy {
        FieldPatch::Replace(payload) => validate_reference(payload, 3, budget)?,
        FieldPatch::Preserve | FieldPatch::Clear | FieldPatch::ReferenceList { .. } => {},
    }

    validate_patch_budget(write.identifier, budget)?;
    match write.identifier {
        FieldPatch::Replace(payload) => {
            str::from_utf8(payload)
                .map(|_| ())
                .map_err(|_| DecodeError::invalid("replacement theme identifier is not UTF-8"))?;
        },
        FieldPatch::Preserve | FieldPatch::Clear | FieldPatch::ReferenceList { .. } => {},
    }

    validate_patch_budget(write.document, budget)?;
    match write.document {
        FieldPatch::Replace(payload) => validate_reference(payload, 3, budget)?,
        FieldPatch::Preserve | FieldPatch::Clear | FieldPatch::ReferenceList { .. } => {},
    }

    for (index, patch) in write.extensions.iter().copied().enumerate() {
        validate_patch_budget(patch, budget)?;
        match patch {
            FieldPatch::Replace(payload) => {
                validate_theme_extension(payload, EXTENSION_FIELDS[index], 3, budget)?;
            },
            FieldPatch::ReferenceList {
                extension,
                field_number,
                references,
            } => {
                if extension_patch_index(extension) != index {
                    return Err(DecodeError::invalid(
                        "theme reference-list rewrite targets the wrong extension",
                    ));
                }
                validate_reference_list_patch(extension, field_number, references, budget)?;
            },
            FieldPatch::Preserve | FieldPatch::Clear => {},
        }
    }
    Ok(())
}

fn validate_reference_list_patch(
    extension: ThemeExtension,
    field_number: u32,
    references: &[ThemeReferenceSnapshot],
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    if !extension_reference_field(extension, field_number) {
        return Err(DecodeError::invalid(format!(
            "field {field_number} is not a TSP.Reference list in this theme extension"
        )));
    }
    for reference in references {
        if reference.identifier == 0 {
            return Err(DecodeError::invalid(
                "TSP.Reference.identifier must be non-zero",
            ));
        }
        let encoded = encoded_reference_field_len(field_number, *reference);
        budget.field()?;
        budget.charge_work(encoded)?;
    }
    Ok(())
}

fn validate_patch_budget(patch: FieldPatch<'_>, budget: &mut Budget) -> Result<(), DecodeError> {
    if let FieldPatch::Replace(payload) = patch {
        if payload.len() > budget.options.max_output_bytes {
            return Err(DecodeError::resource(
                DecodeLimit::OutputBytes,
                payload.len(),
                budget.options.max_output_bytes,
            ));
        }
        budget.charge_work(payload.len())?;
    }
    Ok(())
}

fn measure_rewrite(
    source: &[u8],
    preflight: Preflight<'_>,
    write: ThemeArchiveWrite<'_>,
    budget: &mut Budget,
) -> Result<(usize, usize), DecodeError> {
    let inner_before = budget.fields;
    let inner_bytes = measure_inner(preflight.theme, write, budget)?;
    let inner_fields = budget.fields.saturating_sub(inner_before);
    let maximum = budget.options.max_output_bytes;
    let mut output_bytes = 0usize;
    scan_fields(source, 1, None, budget, |field| {
        if field.number == THEME_SUPER_FIELD {
            output_bytes = output_bytes
                .checked_add(encoded_field_len(THEME_SUPER_FIELD, inner_bytes))
                .ok_or_else(|| output_overflow(maximum))?;
        } else {
            output_bytes = output_bytes
                .checked_add(field.span.end - field.span.start)
                .ok_or_else(|| output_overflow(maximum))?;
        }
        Ok(())
    })?;
    let _ = inner_fields;
    Ok((inner_bytes, output_bytes))
}

fn measure_inner(
    theme: &[u8],
    write: ThemeArchiveWrite<'_>,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    let maximum = budget.options.max_output_bytes;
    let mut output_bytes = 0usize;
    let mut emitted = [false; 8];
    let mut pending_reference_list = None;
    scan_fields(theme, 2, None, budget, |field| {
        if let Some((index, patch)) = patch_for(write, field.number) {
            match patch {
                FieldPatch::Preserve => {
                    output_bytes = output_bytes
                        .checked_add(field.span.end - field.span.start)
                        .ok_or_else(|| output_overflow(maximum))?;
                },
                FieldPatch::Clear => {},
                FieldPatch::Replace(payload) => {
                    if !emitted[index] {
                        output_bytes = output_bytes
                            .checked_add(encoded_field_len(field.number, payload.len()))
                            .ok_or_else(|| output_overflow(maximum))?;
                        emitted[index] = true;
                    }
                },
                FieldPatch::ReferenceList { .. } => {
                    if !emitted[index] {
                        pending_reference_list = Some((field.number, field.span, patch));
                        emitted[index] = true;
                    }
                },
            }
        } else {
            output_bytes = output_bytes
                .checked_add(field.span.end - field.span.start)
                .ok_or_else(|| output_overflow(maximum))?;
        }
        Ok(())
    })?;
    if let Some((
        outer_field_number,
        span,
        FieldPatch::ReferenceList {
            field_number,
            references,
            ..
        },
    )) = pending_reference_list
    {
        let payload_bytes = measure_reference_list_payload(
            span.payload(theme),
            field_number,
            references,
            Some(budget),
        )?;
        if payload_bytes != 0 {
            output_bytes = output_bytes
                .checked_add(encoded_field_len(outer_field_number, payload_bytes))
                .ok_or_else(|| output_overflow(maximum))?;
        }
    }
    for (index, patch) in all_patches(write).iter().copied().enumerate() {
        if !emitted[index] {
            match patch {
                FieldPatch::Replace(payload) => {
                    output_bytes = output_bytes
                        .checked_add(encoded_field_len(patch_field(index), payload.len()))
                        .ok_or_else(|| output_overflow(maximum))?;
                },
                FieldPatch::ReferenceList {
                    field_number,
                    references,
                    ..
                } => {
                    let payload_bytes = measure_reference_list_payload(
                        &[],
                        field_number,
                        references,
                        Some(budget),
                    )?;
                    if payload_bytes != 0 {
                        output_bytes = output_bytes
                            .checked_add(encoded_field_len(patch_field(index), payload_bytes))
                            .ok_or_else(|| output_overflow(maximum))?;
                    }
                },
                FieldPatch::Preserve | FieldPatch::Clear => {},
            }
        }
    }
    Ok(output_bytes)
}

/// Account the linear field traversals performed by the allocation-free
/// emitter. Unknown groups remain one source span here; strict nested
/// validation is charged separately by the candidate readback pass.
fn account_emission(
    wrapper: &[u8],
    theme: &[u8],
    write: ThemeArchiveWrite<'_>,
    output_bytes: usize,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    account_linear_fields(theme, budget)?;
    account_linear_fields(wrapper, budget)?;
    account_linear_fields(theme, budget)?;
    account_emission_reference_walks(theme, write, budget)?;
    budget.charge_work(output_bytes)
}

/// Account the repeated reference-list walks performed by the allocation-free
/// emitter. `emit_wrapper` measures the inner message once, then `emit_inner`
/// measures and emits it again; all three passes use the duplicate/reorder
/// matcher even though their runtime budget parameter is deliberately absent.
fn account_emission_reference_walks(
    theme: &[u8],
    write: ThemeArchiveWrite<'_>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let mut emitted = [false; 8];
    let mut cursor = 0usize;
    while cursor < theme.len() {
        let field = next_field(theme, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme source changed after preflight"))?;
        if let Some((
            index,
            FieldPatch::ReferenceList {
                field_number,
                references,
                ..
            },
        )) = patch_for(write, field.number)
        {
            if emitted[index] {
                continue;
            }
            emitted[index] = true;
            let source = &theme[field.span.payload_start..field.span.payload_end];
            // One sizing pass belongs to the wrapper measurement and one to
            // `emit_inner`'s own measurement. The final emitter walk runs
            // only when that second measurement produces a field payload.
            let first =
                measure_reference_list_payload(source, field_number, references, Some(budget))?;
            let second =
                measure_reference_list_payload(source, field_number, references, Some(budget))?;
            if second != first {
                return Err(DecodeError::projection());
            }
            if second != 0 {
                let third =
                    measure_reference_list_payload(source, field_number, references, Some(budget))?;
                if third != second {
                    return Err(DecodeError::projection());
                }
            }
        }
    }
    for (index, patch) in all_patches(write).iter().copied().enumerate() {
        if emitted[index] {
            continue;
        }
        if let FieldPatch::ReferenceList {
            field_number,
            references,
            ..
        } = patch
        {
            let first =
                measure_reference_list_payload(&[], field_number, references, Some(budget))?;
            let second =
                measure_reference_list_payload(&[], field_number, references, Some(budget))?;
            if second != first {
                return Err(DecodeError::projection());
            }
            if second != 0 {
                let third =
                    measure_reference_list_payload(&[], field_number, references, Some(budget))?;
                if third != second {
                    return Err(DecodeError::projection());
                }
            }
        }
    }
    Ok(())
}

fn account_linear_fields(source: &[u8], budget: &mut Budget) -> Result<(), DecodeError> {
    let mut cursor = 0usize;
    while cursor < source.len() {
        let field = next_field(source, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme source changed after preflight"))?;
        budget.field()?;
        budget.charge_work(field.span.end - field.span.start)?;
    }
    Ok(())
}

/// Precharge strict candidate readback without materializing the candidate.
/// Replacements are represented by encoded lengths and validated through the
/// same bounded nested passes used by ingress.
fn account_candidate_readback(
    source: &[u8],
    preflight: Preflight<'_>,
    write: ThemeArchiveWrite<'_>,
    inner_bytes: usize,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    account_candidate_wrapper(source, inner_bytes, budget)?;
    account_candidate_base(preflight.theme, write, budget)
}

fn account_candidate_wrapper(
    source: &[u8],
    inner_bytes: usize,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let mut cursor = 0usize;
    while cursor < source.len() {
        let field = next_field(source, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme source changed after preflight"))?;
        if field.number == THEME_SUPER_FIELD {
            budget.field()?;
            budget.charge_work(encoded_field_len(THEME_SUPER_FIELD, inner_bytes))?;
        } else {
            account_preserved_field(&source[field.span.start..field.span.end], 1, budget)?;
        }
    }
    Ok(())
}

fn account_candidate_base(
    theme: &[u8],
    write: ThemeArchiveWrite<'_>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let mut present = [false; 8];
    let mut cursor = 0usize;
    while cursor < theme.len() {
        let field = next_field(theme, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme source changed after preflight"))?;
        if let Some((index, patch)) = patch_for(write, field.number) {
            present[index] = true;
            match patch {
                FieldPatch::Preserve => {
                    account_preserved_field(&theme[field.span.start..field.span.end], 2, budget)?;
                    validate_candidate_nested(
                        field.number,
                        &theme[field.span.payload_start..field.span.payload_end],
                        3,
                        budget,
                    )?;
                },
                FieldPatch::Clear => {},
                FieldPatch::Replace(payload) => {
                    account_replacement_field(field.number, payload, budget)?;
                    validate_candidate_nested(field.number, payload, 3, budget)?;
                },
                FieldPatch::ReferenceList {
                    extension,
                    field_number,
                    references,
                } => {
                    account_reference_list_candidate(
                        field.number,
                        extension,
                        field.span.payload(theme),
                        field_number,
                        references,
                        budget,
                    )?;
                },
            }
        } else {
            account_preserved_field(&theme[field.span.start..field.span.end], 2, budget)?;
        }
    }
    for (index, patch) in all_patches(write).iter().copied().enumerate() {
        if present[index] {
            continue;
        }
        match patch {
            FieldPatch::Replace(payload) => {
                let number = patch_field(index);
                account_replacement_field(number, payload, budget)?;
                validate_candidate_nested(number, payload, 3, budget)?;
            },
            FieldPatch::ReferenceList {
                extension,
                field_number,
                references,
            } => {
                account_reference_list_candidate(
                    patch_field(index),
                    extension,
                    &[],
                    field_number,
                    references,
                    budget,
                )?;
            },
            FieldPatch::Preserve | FieldPatch::Clear => {},
        }
    }
    Ok(())
}

fn account_replacement_field(
    number: u32,
    payload: &[u8],
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.field()?;
    budget.charge_work(encoded_field_len(number, payload.len()))
}

fn account_reference_list_candidate(
    outer_field_number: u32,
    extension: ThemeExtension,
    source: &[u8],
    field_number: u32,
    references: &[ThemeReferenceSnapshot],
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let payload_bytes =
        measure_reference_list_payload(source, field_number, references, Some(budget))?;
    if payload_bytes == 0 {
        return Ok(());
    }
    budget.field()?;
    budget.charge_work(encoded_field_len(outer_field_number, payload_bytes))?;
    // Unknown and retained source records receive the same strict validation
    // pass as an ordinary extension replacement. Canonical replacement
    // records are charged separately below without allocating a temporary
    // encoded payload.
    validate_theme_extension(source, extension_field(extension), 3, budget)?;
    for reference in references {
        account_reference_snapshot(*reference, 4, budget)?;
    }
    Ok(())
}

fn account_reference_snapshot(
    reference: ThemeReferenceSnapshot,
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.observe_depth(depth)?;
    let identifier = encoded_varint_field_len(1, reference.identifier);
    budget.field()?;
    budget.charge_work(identifier)?;
    if let Some(value) = reference.deprecated_type {
        budget.field()?;
        budget.charge_work(encoded_varint_field_len(2, value as i64 as u64))?;
    }
    if let Some(value) = reference.deprecated_is_external {
        budget.field()?;
        budget.charge_work(encoded_varint_field_len(3, u64::from(value)))?;
    }
    Ok(())
}

fn account_preserved_field(
    source: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let mut cursor = 0usize;
    let field = next_field(source, &mut cursor)
        .ok_or_else(|| DecodeError::invalid("theme source changed after preflight"))?;
    if field.wire_type == 3 {
        // Strict ingress recursively counts fields inside unknown groups.
        scan_fields(source, depth, None, budget, |_| Ok(()))?;
    } else {
        budget.field()?;
        budget.charge_work(source.len())?;
    }
    Ok(())
}

fn validate_candidate_nested(
    number: u32,
    payload: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    match number {
        LEGACY_STYLESHEET_FIELD | DOCUMENT_STYLESHEET_FIELD => {
            validate_reference(payload, depth, budget)
        },
        OLD_PRESET_UUIDS_FIELD | NEW_PRESET_UUIDS_FIELD => validate_uuid(payload, depth, budget),
        COLOR_PRESETS_FIELD => validate_color(payload, depth, budget),
        DRAWING_PRESETS_FIELD
        | TEXT_PRESETS_FIELD
        | CHART_PRESETS_FIELD
        | TABLE_PRESETS_FIELD
        | APPLICATION_PRESETS_FIELD => validate_theme_extension(payload, number, depth, budget),
        _ => Ok(()),
    }
}

fn emit_wrapper(
    source: &[u8],
    preflight: Preflight<'_>,
    write: ThemeArchiveWrite<'_>,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let inner_bytes = measure_inner_without_budget(preflight.theme, write)?;
    let mut cursor = 0usize;
    while cursor < source.len() {
        let field = next_field(source, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme source changed after preflight"))?;
        if field.number == THEME_SUPER_FIELD {
            push_varint((u64::from(THEME_SUPER_FIELD) << 3) | 2, output);
            push_varint(inner_bytes as u64, output);
            emit_inner(preflight.theme, write, output)?;
        } else {
            output.extend_from_slice(&source[field.span.start..field.span.end]);
        }
    }
    Ok(())
}

fn emit_inner(
    theme: &[u8],
    write: ThemeArchiveWrite<'_>,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut emitted = [false; 8];
    let mut cursor = 0usize;
    while cursor < theme.len() {
        let field = next_field(theme, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme source changed after preflight"))?;
        if let Some((index, patch)) = patch_for(write, field.number) {
            match patch {
                FieldPatch::Preserve => {
                    output.extend_from_slice(&theme[field.span.start..field.span.end]);
                },
                FieldPatch::Clear => {},
                FieldPatch::Replace(payload) => {
                    if !emitted[index] {
                        append_raw_field(output, field.number, payload);
                        emitted[index] = true;
                    }
                },
                FieldPatch::ReferenceList {
                    field_number,
                    references,
                    ..
                } => {
                    if !emitted[index] {
                        let source = &theme[field.span.payload_start..field.span.payload_end];
                        let payload_bytes =
                            measure_reference_list_payload(source, field_number, references, None)?;
                        if payload_bytes != 0 {
                            push_varint((u64::from(field.number) << 3) | 2, output);
                            push_varint(payload_bytes as u64, output);
                            emit_reference_list_payload(source, field_number, references, output)?;
                        }
                        emitted[index] = true;
                    }
                },
            }
        } else {
            output.extend_from_slice(&theme[field.span.start..field.span.end]);
        }
    }
    for (index, patch) in all_patches(write).iter().copied().enumerate() {
        if !emitted[index] {
            match patch {
                FieldPatch::Replace(payload) => {
                    append_raw_field(output, patch_field(index), payload);
                },
                FieldPatch::ReferenceList {
                    field_number,
                    references,
                    ..
                } => {
                    let payload_bytes =
                        measure_reference_list_payload(&[], field_number, references, None)?;
                    if payload_bytes != 0 {
                        push_varint((u64::from(patch_field(index)) << 3) | 2, output);
                        push_varint(payload_bytes as u64, output);
                        emit_reference_list_payload(&[], field_number, references, output)?;
                    }
                },
                FieldPatch::Preserve | FieldPatch::Clear => {},
            }
        }
    }
    Ok(())
}

fn measure_inner_without_budget(
    theme: &[u8],
    write: ThemeArchiveWrite<'_>,
) -> Result<usize, DecodeError> {
    let mut output_bytes = 0usize;
    let mut emitted = [false; 8];
    let mut cursor = 0usize;
    while cursor < theme.len() {
        let field = next_field(theme, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme source changed after preflight"))?;
        if let Some((index, patch)) = patch_for(write, field.number) {
            match patch {
                FieldPatch::Preserve => {
                    output_bytes = output_bytes
                        .checked_add(field.span.end - field.span.start)
                        .ok_or_else(|| DecodeError::invalid("theme output length overflow"))?;
                },
                FieldPatch::Clear => {},
                FieldPatch::Replace(payload) if !emitted[index] => {
                    output_bytes = output_bytes
                        .checked_add(encoded_field_len(field.number, payload.len()))
                        .ok_or_else(|| DecodeError::invalid("theme output length overflow"))?;
                    emitted[index] = true;
                },
                FieldPatch::Replace(_) => {},
                FieldPatch::ReferenceList {
                    field_number,
                    references,
                    ..
                } if !emitted[index] => {
                    let source = &theme[field.span.payload_start..field.span.payload_end];
                    let payload_bytes =
                        measure_reference_list_payload(source, field_number, references, None)?;
                    if payload_bytes != 0 {
                        output_bytes = output_bytes
                            .checked_add(encoded_field_len(field.number, payload_bytes))
                            .ok_or_else(|| DecodeError::invalid("theme output length overflow"))?;
                    }
                    emitted[index] = true;
                },
                FieldPatch::ReferenceList { .. } => {},
            }
        } else {
            output_bytes = output_bytes
                .checked_add(field.span.end - field.span.start)
                .ok_or_else(|| DecodeError::invalid("theme output length overflow"))?;
        }
    }
    for (index, patch) in all_patches(write).iter().copied().enumerate() {
        if !emitted[index] {
            match patch {
                FieldPatch::Replace(payload) => {
                    output_bytes = output_bytes
                        .checked_add(encoded_field_len(patch_field(index), payload.len()))
                        .ok_or_else(|| DecodeError::invalid("theme output length overflow"))?;
                },
                FieldPatch::ReferenceList {
                    field_number,
                    references,
                    ..
                } => {
                    let payload_bytes =
                        measure_reference_list_payload(&[], field_number, references, None)?;
                    if payload_bytes != 0 {
                        output_bytes = output_bytes
                            .checked_add(encoded_field_len(patch_field(index), payload_bytes))
                            .ok_or_else(|| DecodeError::invalid("theme output length overflow"))?;
                    }
                },
                FieldPatch::Preserve | FieldPatch::Clear => {},
            }
        }
    }
    Ok(output_bytes)
}

/// Account the `ReferenceFields` replay used by semantic no-op verification.
/// The strict source pass has already validated the extension, so this helper
/// only charges the additional outer and nested scans plus scalar comparisons
/// performed by the borrowed iterator.
fn account_reference_fields_match(
    payload: Option<&[u8]>,
    field_number: u32,
    references: &[ThemeReferenceSnapshot],
    budget: &mut Budget,
) -> Result<bool, DecodeError> {
    let Some(payload) = payload else {
        budget.charge_work(
            references
                .len()
                .checked_mul(size_of::<ThemeReferenceSnapshot>())
                .ok_or_else(|| work_overflow(usize::MAX))?,
        )?;
        return Ok(references.is_empty());
    };

    // ReferenceFields reparses the extension envelope and then reparses each
    // selected TSP.Reference record while producing its semantic snapshot.
    account_linear_fields(payload, budget)?;
    let mut cursor = 0usize;
    while cursor < payload.len() {
        let field = next_field(payload, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme extension source changed after match"))?;
        if field.number == field_number {
            scan_fields(field.span.payload(payload), 4, None, budget, |_| Ok(()))?;
        }
    }
    budget.charge_work(
        references
            .len()
            .checked_mul(size_of::<ThemeReferenceSnapshot>())
            .ok_or_else(|| work_overflow(usize::MAX))?,
    )?;
    Ok(ReferenceFields::new(payload, field_number).eq(references.iter().copied()))
}

fn account_matches_write(
    snapshot: ThemeArchiveSnapshot<'_>,
    write: ThemeArchiveWrite<'_>,
    budget: &mut Budget,
) -> Result<bool, DecodeError> {
    let mut matches = account_patch_match(snapshot.legacy_stylesheet(), write.legacy, budget)?;
    matches &= account_patch_match(
        snapshot.theme_identifier().map(str::as_bytes),
        write.identifier,
        budget,
    )?;
    matches &= account_patch_match(snapshot.document_stylesheet(), write.document, budget)?;
    for (index, patch) in write.extensions.iter().copied().enumerate() {
        let current = snapshot_extension_payload(
            snapshot,
            extension_from_field(EXTENSION_FIELDS[index]).expect("known extension field"),
        );
        match patch {
            FieldPatch::ReferenceList {
                field_number,
                references,
                ..
            } => {
                matches &=
                    account_reference_fields_match(current, field_number, references, budget)?;
            },
            patch => {
                matches &= account_patch_match(current, patch, budget)?;
            },
        }
    }
    Ok(matches)
}

fn account_patch_match(
    current: Option<&[u8]>,
    patch: FieldPatch<'_>,
    budget: &mut Budget,
) -> Result<bool, DecodeError> {
    match patch {
        FieldPatch::Preserve => Ok(true),
        FieldPatch::Clear => Ok(current.is_none()),
        FieldPatch::Replace(value) => {
            // Slice equality may inspect every replacement byte before the
            // semantic match is decided. The write-validation pass already
            // admits the replacement itself; charge this comparison replay
            // separately.
            budget.charge_work(value.len())?;
            Ok(current == Some(value))
        },
        FieldPatch::ReferenceList { .. } => Ok(false),
    }
}

fn account_patch_comparison(patch: FieldPatch<'_>, budget: &mut Budget) -> Result<(), DecodeError> {
    if let FieldPatch::Replace(value) = patch {
        budget.charge_work(value.len())?;
    }
    Ok(())
}

/// Precharge the semantic match against the candidate extension payload. The
/// source extension is an admitted upper bound for preserved unknown fields;
/// requested references account for records added or reordered by the write.
fn account_candidate_matches_write(
    theme: &[u8],
    write: ThemeArchiveWrite<'_>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    account_patch_comparison(write.legacy, budget)?;
    account_patch_comparison(write.identifier, budget)?;
    account_patch_comparison(write.document, budget)?;
    for patch in write.extensions.iter().copied() {
        account_patch_comparison(patch, budget)?;
    }
    let mut present = [false; 5];
    let mut cursor = 0usize;
    while cursor < theme.len() {
        let field = next_field(theme, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme source changed after match admission"))?;
        let Some(index) = extension_index(field.number) else {
            continue;
        };
        let FieldPatch::ReferenceList {
            field_number,
            references,
            ..
        } = write.extensions[index]
        else {
            continue;
        };
        if present[index] {
            continue;
        }
        present[index] = true;
        let source = field.span.payload(theme);
        validate_theme_extension(source, field.number, 3, budget)?;
        account_linear_fields(source, budget)?;
        for reference in references {
            budget.field()?;
            budget.charge_work(encoded_reference_field_len(field_number, *reference))?;
            account_reference_snapshot(*reference, 4, budget)?;
        }
    }
    for (index, patch) in write.extensions.iter().copied().enumerate() {
        if present[index] {
            continue;
        }
        if let FieldPatch::ReferenceList {
            field_number,
            references,
            ..
        } = patch
        {
            for reference in references {
                budget.field()?;
                budget.charge_work(encoded_reference_field_len(field_number, *reference))?;
                account_reference_snapshot(*reference, 4, budget)?;
            }
        }
    }
    Ok(())
}

fn snapshot_extension_payload<'source>(
    snapshot: ThemeArchiveSnapshot<'source>,
    extension: ThemeExtension,
) -> Option<&'source [u8]> {
    match extension {
        ThemeExtension::Drawing => snapshot.drawing_presets(),
        ThemeExtension::Text => snapshot.text_presets(),
        ThemeExtension::Chart => snapshot.chart_presets(),
        ThemeExtension::Table => snapshot.table_presets(),
        ThemeExtension::Application => snapshot.application_presets(),
    }
}

fn account_reference_projection<T, I, F>(
    snapshot: ThemeArchiveSnapshot<'_>,
    selectors: I,
    route: &F,
    budget: &mut Budget,
) -> Result<usize, DecodeError>
where
    I: IntoIterator<Item = T>,
    F: Fn(T) -> (ThemeExtension, u32),
{
    let mut count = 0usize;
    for selector in selectors {
        let (extension, field_number) = route(selector);
        // Charge the route itself so a caller cannot submit an arbitrarily
        // large selector slice outside the source-derived field ceiling.
        budget.field()?;
        let Some(payload) = snapshot_extension_payload(snapshot, extension) else {
            continue;
        };
        let mut cursor = 0usize;
        while cursor < payload.len() {
            let field = next_field(payload, &mut cursor)
                .ok_or_else(|| DecodeError::invalid("theme extension source changed after scan"))?;
            budget.field()?;
            budget.charge_work(field.span.end - field.span.start)?;
            if field.number == field_number {
                count = count
                    .checked_add(1)
                    .ok_or_else(|| field_overflow(usize::MAX))?;
                // The projection copies the typed scalar values out of the
                // nested record after the outer field walk.
                budget.charge_work(field.span.payload(payload).len())?;
            }
        }
    }
    Ok(count)
}

fn patch_for(write: ThemeArchiveWrite<'_>, field_number: u32) -> Option<(usize, FieldPatch<'_>)> {
    match field_number {
        LEGACY_STYLESHEET_FIELD => Some((0, write.legacy)),
        THEME_IDENTIFIER_FIELD => Some((1, write.identifier)),
        DOCUMENT_STYLESHEET_FIELD => Some((2, write.document)),
        DRAWING_PRESETS_FIELD => Some((3, write.extensions[0])),
        TEXT_PRESETS_FIELD => Some((4, write.extensions[1])),
        CHART_PRESETS_FIELD => Some((5, write.extensions[2])),
        TABLE_PRESETS_FIELD => Some((6, write.extensions[3])),
        APPLICATION_PRESETS_FIELD => Some((7, write.extensions[4])),
        _ => None,
    }
}

fn all_patches(write: ThemeArchiveWrite<'_>) -> [FieldPatch<'_>; 8] {
    [
        write.legacy,
        write.identifier,
        write.document,
        write.extensions[0],
        write.extensions[1],
        write.extensions[2],
        write.extensions[3],
        write.extensions[4],
    ]
}

fn patch_field(index: usize) -> u32 {
    match index {
        0 => LEGACY_STYLESHEET_FIELD,
        1 => THEME_IDENTIFIER_FIELD,
        2 => DOCUMENT_STYLESHEET_FIELD,
        index => EXTENSION_FIELDS[index - 3],
    }
}

fn ensure_requirements(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    let checks = [
        (
            DecodeLimit::OutputBytes,
            requirements.output_bytes,
            limits.output_bytes,
        ),
        (DecodeLimit::Fields, requirements.fields, limits.fields),
        (
            DecodeLimit::WorkBytes,
            requirements.work_bytes,
            limits.work_bytes,
        ),
        (
            DecodeLimit::Nesting,
            requirements.max_depth as usize,
            limits.max_depth as usize,
        ),
        (
            DecodeLimit::Allocations,
            requirements.allocations,
            limits.allocations,
        ),
        (
            DecodeLimit::RetainedBytes,
            requirements.retained_bytes,
            limits.retained_bytes,
        ),
        (
            DecodeLimit::ScratchBytes,
            requirements.scratch_bytes,
            limits.scratch_bytes,
        ),
    ];
    for (limit, observed, maximum) in checks {
        if observed > maximum {
            return Err(DecodeError::resource(limit, observed, maximum));
        }
    }
    Ok(())
}

#[derive(Debug)]
struct Budget {
    options: DecodeOptions,
    source_bytes: usize,
    source_token: SourceToken,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Self {
        Self {
            options,
            source_bytes: source.len(),
            source_token: SourceToken::for_source(source),
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
        }
    }

    fn from_report(options: DecodeOptions, report: DecodeReport) -> Self {
        Self {
            options,
            source_bytes: report.source_bytes,
            source_token: report.source_token,
            fields: report.fields,
            work_bytes: report.work_bytes,
            max_depth: report.max_depth,
        }
    }

    fn add_report(&mut self, report: DecodeReport) -> Result<(), DecodeError> {
        self.fields = self
            .fields
            .checked_add(report.fields)
            .ok_or_else(|| field_overflow(self.options.max_fields))?;
        self.work_bytes = self
            .work_bytes
            .checked_add(report.work_bytes)
            .ok_or_else(|| work_overflow(self.options.max_work_bytes))?;
        self.max_depth = self.max_depth.max(report.max_depth);
        if self.fields > self.options.max_fields {
            return Err(DecodeError::resource(
                DecodeLimit::Fields,
                self.fields,
                self.options.max_fields,
            ));
        }
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::resource(
                DecodeLimit::WorkBytes,
                self.work_bytes,
                self.options.max_work_bytes,
            ));
        }
        if self.max_depth > self.options.recursion_limit {
            return Err(DecodeError::resource(
                DecodeLimit::Nesting,
                self.max_depth as usize,
                self.options.recursion_limit as usize,
            ));
        }
        Ok(())
    }

    fn field(&mut self) -> Result<(), DecodeError> {
        self.fields = self
            .fields
            .checked_add(1)
            .ok_or_else(|| field_overflow(self.options.max_fields))?;
        if self.fields > self.options.max_fields {
            return Err(DecodeError::resource(
                DecodeLimit::Fields,
                self.fields,
                self.options.max_fields,
            ));
        }
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), DecodeError> {
        self.work_bytes = self
            .work_bytes
            .checked_add(amount)
            .ok_or_else(|| work_overflow(self.options.max_work_bytes))?;
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::resource(
                DecodeLimit::WorkBytes,
                self.work_bytes,
                self.options.max_work_bytes,
            ));
        }
        Ok(())
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        self.max_depth = self.max_depth.max(depth);
        if depth > self.options.recursion_limit {
            return Err(DecodeError::resource(
                DecodeLimit::Nesting,
                depth as usize,
                self.options.recursion_limit as usize,
            ));
        }
        Ok(())
    }

    fn buffa(&mut self, bytes: usize) -> Result<(), DecodeError> {
        self.charge_work(bytes)
    }

    fn decode_report(&self) -> DecodeReport {
        DecodeReport {
            source_bytes: self.source_bytes,
            source_token: self.source_token,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            allocations: 0,
            retained_bytes: self.source_bytes,
            scratch_bytes: 0,
        }
    }
}

fn validate_options(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    if options.max_message_bytes > MAX_AUTOMATIC_LIMIT {
        return Err(DecodeError::resource(
            DecodeLimit::InputBytes,
            options.max_message_bytes,
            MAX_AUTOMATIC_LIMIT,
        ));
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::InputBytes,
            source.len(),
            options.max_message_bytes,
        ));
    }
    for (limit, value) in [
        (DecodeLimit::Fields, options.max_fields),
        (DecodeLimit::WorkBytes, options.max_work_bytes),
        (DecodeLimit::OutputBytes, options.max_output_bytes),
        (DecodeLimit::Allocations, options.max_allocations),
        (DecodeLimit::RetainedBytes, options.max_retained_bytes),
        (DecodeLimit::ScratchBytes, options.max_scratch_bytes),
    ] {
        if value > MAX_AUTOMATIC_LIMIT {
            return Err(DecodeError::resource(limit, value, MAX_AUTOMATIC_LIMIT));
        }
    }
    if source.len() > options.max_retained_bytes {
        return Err(DecodeError::resource(
            DecodeLimit::RetainedBytes,
            source.len(),
            options.max_retained_bytes,
        ));
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

fn scan_fields<F>(
    source: &[u8],
    depth: u32,
    expected_end_group: Option<u32>,
    budget: &mut Budget,
    visit: F,
) -> Result<usize, DecodeError>
where
    F: FnMut(WireField) -> Result<(), DecodeError>,
{
    let mut visit = visit;
    scan_fields_dyn(source, depth, expected_end_group, budget, &mut visit)
}

fn scan_fields_dyn(
    source: &[u8],
    depth: u32,
    expected_end_group: Option<u32>,
    budget: &mut Budget,
    visit: &mut dyn FnMut(WireField) -> Result<(), DecodeError>,
) -> Result<usize, DecodeError> {
    budget.observe_depth(depth)?;
    let mut cursor = 0usize;
    while cursor < source.len() {
        let start = cursor;
        let (tag, after_tag) = read_varint(source, cursor)?;
        cursor = after_tag;
        let number = u32::try_from(tag >> 3)
            .map_err(|_| DecodeError::invalid("protobuf field number overflows u32"))?;
        let wire_type = u8::try_from(tag & 7)
            .map_err(|_| DecodeError::invalid("protobuf wire type overflows u8"))?;
        if number == 0 || number > 0x1fff_ffff {
            return Err(DecodeError::invalid(
                "protobuf field number is out of range",
            ));
        }
        budget.field()?;
        if wire_type == 4 {
            if expected_end_group == Some(number) {
                budget.charge_work(cursor - start)?;
                return Ok(cursor);
            }
            return Err(DecodeError::invalid("unexpected protobuf end-group"));
        }
        let mut payload_start = cursor;
        match wire_type {
            0 => {
                let (_, after_value) = read_varint(source, cursor)?;
                cursor = after_value;
            },
            1 => {
                cursor = cursor
                    .checked_add(8)
                    .ok_or_else(|| DecodeError::invalid("fixed64 length overflow"))?;
                if cursor > source.len() {
                    return Err(DecodeError::invalid("truncated protobuf fixed64"));
                }
            },
            2 => {
                let (length, after_length) = read_varint(source, cursor)?;
                let length = usize::try_from(length)
                    .map_err(|_| DecodeError::invalid("length-delimited field is too large"))?;
                payload_start = after_length;
                cursor = after_length
                    .checked_add(length)
                    .ok_or_else(|| DecodeError::invalid("length-delimited field overflows"))?;
                if cursor > source.len() {
                    return Err(DecodeError::invalid("truncated length-delimited field"));
                }
            },
            3 => {
                let nested = &source[cursor..];
                let consumed = scan_fields_dyn(
                    nested,
                    depth + 1,
                    Some(number),
                    budget,
                    &mut |_: WireField| Ok(()),
                )?;
                cursor = cursor
                    .checked_add(consumed)
                    .ok_or_else(|| DecodeError::invalid("group length overflows"))?;
            },
            5 => {
                cursor = cursor
                    .checked_add(4)
                    .ok_or_else(|| DecodeError::invalid("fixed32 length overflow"))?;
                if cursor > source.len() {
                    return Err(DecodeError::invalid("truncated protobuf fixed32"));
                }
            },
            _ => return Err(DecodeError::invalid("invalid protobuf wire type")),
        }
        budget.charge_work(cursor - start)?;
        visit(WireField {
            number,
            wire_type,
            span: Span {
                start,
                payload_start,
                payload_end: cursor,
                end: cursor,
            },
        })?;
    }
    if expected_end_group.is_some() {
        Err(DecodeError::invalid("unterminated protobuf group"))
    } else {
        Ok(cursor)
    }
}

fn validate_theme_extension(
    source: &[u8],
    extension_field: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let extension = extension_from_field(extension_field)
        .ok_or_else(|| DecodeError::invalid("unknown TSS.ThemeArchive extension"))?;

    // The complete extension remains source-preserving and future fields are
    // accepted. Known repeated TSP.Reference lists still receive a strict
    // nested validation pass so callers can safely project their identifiers.
    scan_fields(source, depth, None, budget, |field| {
        if extension_reference_field(extension, field.number) {
            require_wire(field, 2, "theme preset reference")?;
        }
        Ok(())
    })?;

    let mut cursor = 0usize;
    while cursor < source.len() {
        let field = next_field(source, &mut cursor)
            .ok_or_else(|| DecodeError::invalid("theme extension source changed after scan"))?;
        if extension_reference_field(extension, field.number) {
            validate_reference(
                &source[field.span.payload_start..field.span.payload_end],
                depth + 1,
                budget,
            )?;
        }
    }
    Ok(())
}

fn validate_reference(source: &[u8], depth: u32, budget: &mut Budget) -> Result<(), DecodeError> {
    let mut identifier = false;
    let mut deprecated_type = false;
    let mut deprecated_is_external = false;
    scan_fields(source, depth, None, budget, |field| {
        match field.number {
            1 => {
                if identifier {
                    return Err(DecodeError::duplicate("TSP.Reference.identifier"));
                }
                identifier = true;
                require_wire(field, 0, "TSP.Reference.identifier")?;
                if scalar_value(source, field)? == 0 {
                    return Err(DecodeError::invalid(
                        "TSP.Reference.identifier must be non-zero",
                    ));
                }
            },
            2 => {
                if deprecated_type {
                    return Err(DecodeError::duplicate("TSP.Reference.deprecated_type"));
                }
                deprecated_type = true;
                require_wire(field, 0, "TSP.Reference.deprecated_type")?;
                require_canonical_int32(
                    scalar_value(source, field)?,
                    "TSP.Reference.deprecated_type",
                )?;
            },
            3 => {
                if deprecated_is_external {
                    return Err(DecodeError::duplicate(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                deprecated_is_external = true;
                require_wire(field, 0, "TSP.Reference.deprecated_is_external")?;
                require_canonical_bool(
                    scalar_value(source, field)?,
                    "TSP.Reference.deprecated_is_external",
                )?;
            },
            _ => {},
        }
        Ok(())
    })?;
    if identifier {
        Ok(())
    } else {
        Err(DecodeError::missing("TSP.Reference.identifier"))
    }
}

fn validate_uuid(source: &[u8], depth: u32, budget: &mut Budget) -> Result<(), DecodeError> {
    let mut lower = false;
    let mut upper = false;
    scan_fields(source, depth, None, budget, |field| {
        match field.number {
            1 => {
                if lower {
                    return Err(DecodeError::duplicate("TSP.UUID.lower"));
                }
                require_wire(field, 0, "TSP.UUID.lower")?;
                lower = true;
            },
            2 => {
                if upper {
                    return Err(DecodeError::duplicate("TSP.UUID.upper"));
                }
                require_wire(field, 0, "TSP.UUID.upper")?;
                upper = true;
            },
            _ => {},
        }
        Ok(())
    })?;
    if lower && upper {
        Ok(())
    } else {
        Err(DecodeError::invalid(
            "TSP.UUID.lower and upper are required",
        ))
    }
}

fn validate_color(source: &[u8], depth: u32, budget: &mut Budget) -> Result<(), DecodeError> {
    let mut seen = [false; 11];
    scan_fields(source, depth, None, budget, |field| {
        match field.number {
            1 => {
                if seen[0] {
                    return Err(DecodeError::duplicate("TSP.Color.model"));
                }
                require_wire(field, 0, "TSP.Color.model")?;
                require_canonical_int32(scalar_value(source, field)?, "TSP.Color.model")?;
                seen[0] = true;
            },
            3..=11 => {
                let index = usize::try_from(field.number - 2).unwrap_or(usize::MAX);
                if seen.get(index).copied().unwrap_or(false) {
                    return Err(DecodeError::duplicate(format!(
                        "TSP.Color field {}",
                        field.number
                    )));
                }
                require_wire(field, 5, "TSP.Color component")?;
                if let Some(value) = seen.get_mut(index) {
                    *value = true;
                }
            },
            12 => {
                if seen[10] {
                    return Err(DecodeError::duplicate("TSP.Color.rgbspace"));
                }
                require_wire(field, 0, "TSP.Color.rgbspace")?;
                require_canonical_int32(scalar_value(source, field)?, "TSP.Color.rgbspace")?;
                seen[10] = true;
            },
            _ => {},
        }
        Ok(())
    })?;
    if seen[0] {
        Ok(())
    } else {
        Err(DecodeError::missing("TSP.Color.model"))
    }
}

fn scalar_value(source: &[u8], field: WireField) -> Result<u64, DecodeError> {
    let (value, end) = read_varint(source, field.span.payload_start)?;
    if end != field.span.end {
        return Err(DecodeError::invalid(
            "protobuf scalar payload has trailing bytes",
        ));
    }
    Ok(value)
}

fn require_canonical_bool(value: u64, field: &str) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical(format!(
            "{field} bool scalar is not zero or one"
        ))),
    }
}

fn require_canonical_int32(value: u64, field: &str) -> Result<i32, DecodeError> {
    const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;
    if value > i32::MAX as u64 && value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical(format!(
            "{field} is not a sign-extended int32"
        )));
    }
    Ok(value as u32 as i32)
}

fn require_wire(field: WireField, expected: u8, name: &str) -> Result<(), DecodeError> {
    if field.wire_type == expected {
        Ok(())
    } else {
        Err(DecodeError::invalid(format!(
            "{name} has wire type {}, expected {expected}",
            field.wire_type
        )))
    }
}

fn read_varint(source: &[u8], start: usize) -> Result<(u64, usize), DecodeError> {
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *source
            .get(start + index)
            .ok_or_else(|| DecodeError::invalid("truncated protobuf varint"))?;
        if index > 0 && byte == 0 {
            return Err(DecodeError::noncanonical(
                "protobuf varint is not canonically encoded",
            ));
        }
        if index == 9 && byte > 1 {
            return Err(DecodeError::invalid("protobuf varint is too long"));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            return Ok((value, start + index + 1));
        }
    }
    Err(DecodeError::invalid("protobuf varint is too long"))
}

fn read_varint_unchecked(source: &[u8], start: usize) -> Option<(u64, usize)> {
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *source.get(start + index)?;
        if index > 0 && byte == 0 {
            return None;
        }
        if index == 9 && byte > 1 {
            return None;
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            return Some((value, start + index + 1));
        }
    }
    None
}

fn next_field(source: &[u8], cursor: &mut usize) -> Option<WireField> {
    let start = *cursor;
    let (tag, after_tag) = read_varint_unchecked(source, start)?;
    let number = u32::try_from(tag >> 3).ok()?;
    let wire_type = u8::try_from(tag & 7).ok()?;
    if number == 0 || number > 0x1fff_ffff || wire_type == 4 || wire_type == 6 || wire_type == 7 {
        return None;
    }
    let mut payload_start = after_tag;
    let end = match wire_type {
        0 => read_varint_unchecked(source, after_tag)?.1,
        1 => after_tag.checked_add(8)?,
        2 => {
            let (length, after_length) = read_varint_unchecked(source, after_tag)?;
            payload_start = after_length;
            after_length.checked_add(usize::try_from(length).ok()?)?
        },
        3 => skip_group_unchecked(source, after_tag, number)?,
        5 => after_tag.checked_add(4)?,
        _ => return None,
    };
    if end > source.len() {
        return None;
    }
    *cursor = end;
    Some(WireField {
        number,
        wire_type,
        span: Span {
            start,
            payload_start,
            payload_end: end,
            end,
        },
    })
}

fn skip_group_unchecked(source: &[u8], mut cursor: usize, expected: u32) -> Option<usize> {
    while cursor < source.len() {
        let (tag, after_tag) = read_varint_unchecked(source, cursor)?;
        cursor = after_tag;
        let number = u32::try_from(tag >> 3).ok()?;
        let wire_type = u8::try_from(tag & 7).ok()?;
        match wire_type {
            0 => cursor = read_varint_unchecked(source, cursor)?.1,
            1 => cursor = cursor.checked_add(8)?,
            2 => {
                let (length, after_length) = read_varint_unchecked(source, cursor)?;
                cursor = after_length.checked_add(usize::try_from(length).ok()?)?;
            },
            3 => cursor = skip_group_unchecked(source, cursor, number)?,
            4 if number == expected => return Some(cursor),
            4 => return None,
            5 => cursor = cursor.checked_add(4)?,
            _ => return None,
        }
        if cursor > source.len() {
            return None;
        }
    }
    None
}

fn payload(source: &[u8], span: Option<Span>) -> Option<&[u8]> {
    span.map(|span| span.payload(source))
}

fn reference_snapshot(source: &[u8]) -> ThemeReferenceSnapshot {
    let mut identifier = 0u64;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    let mut cursor = 0usize;
    while cursor < source.len() {
        let Some(field) = next_field(source, &mut cursor) else {
            break;
        };
        let value =
            read_varint_unchecked(source, field.span.payload_start).map_or(0, |(value, _)| value);
        match field.number {
            1 => identifier = value,
            2 => {
                let value = value as u32;
                deprecated_type = Some(i32::from_ne_bytes(value.to_ne_bytes()));
            },
            3 => deprecated_is_external = Some(value != 0),
            _ => {},
        }
    }
    ThemeReferenceSnapshot {
        identifier,
        deprecated_type,
        deprecated_is_external,
    }
}

fn extension_index(field_number: u32) -> Option<usize> {
    EXTENSION_FIELDS
        .iter()
        .position(|number| *number == field_number)
}

fn extension_from_field(field_number: u32) -> Option<ThemeExtension> {
    match field_number {
        DRAWING_PRESETS_FIELD => Some(ThemeExtension::Drawing),
        TEXT_PRESETS_FIELD => Some(ThemeExtension::Text),
        CHART_PRESETS_FIELD => Some(ThemeExtension::Chart),
        TABLE_PRESETS_FIELD => Some(ThemeExtension::Table),
        APPLICATION_PRESETS_FIELD => Some(ThemeExtension::Application),
        _ => None,
    }
}

fn extension_field(extension: ThemeExtension) -> u32 {
    match extension {
        ThemeExtension::Drawing => DRAWING_PRESETS_FIELD,
        ThemeExtension::Text => TEXT_PRESETS_FIELD,
        ThemeExtension::Chart => CHART_PRESETS_FIELD,
        ThemeExtension::Table => TABLE_PRESETS_FIELD,
        ThemeExtension::Application => APPLICATION_PRESETS_FIELD,
    }
}

fn extension_patch_index(extension: ThemeExtension) -> usize {
    extension_index(extension_field(extension)).expect("known theme extension")
}

fn extension_reference_field(extension: ThemeExtension, field_number: u32) -> bool {
    match extension {
        ThemeExtension::Drawing => (4..=9).contains(&field_number),
        ThemeExtension::Text => (1..=8).contains(&field_number),
        ThemeExtension::Chart => field_number == 1,
        ThemeExtension::Table => matches!(field_number, 1 | 2),
        ThemeExtension::Application => matches!(field_number, 1 | 2),
    }
}

fn is_known_base_or_extension(number: u32) -> bool {
    matches!(
        number,
        LEGACY_STYLESHEET_FIELD
            | THEME_IDENTIFIER_FIELD
            | DOCUMENT_STYLESHEET_FIELD
            | OLD_PRESET_UUIDS_FIELD
            | NEW_PRESET_UUIDS_FIELD
            | COLOR_PRESETS_FIELD
            | DRAWING_PRESETS_FIELD
            | TEXT_PRESETS_FIELD
            | CHART_PRESETS_FIELD
            | TABLE_PRESETS_FIELD
            | APPLICATION_PRESETS_FIELD
    )
}

fn append_raw_field(output: &mut Vec<u8>, field_number: u32, payload: &[u8]) {
    push_varint((u64::from(field_number) << 3) | 2, output);
    push_varint(payload.len() as u64, output);
    output.extend_from_slice(payload);
}

fn append_varint_field(output: &mut Vec<u8>, field_number: u32, value: u64) {
    push_varint(u64::from(field_number) << 3, output);
    push_varint(value, output);
}

fn append_reference_field(
    output: &mut Vec<u8>,
    field_number: u32,
    reference: ThemeReferenceSnapshot,
) {
    let payload_len = reference_payload_len(reference);
    push_varint((u64::from(field_number) << 3) | 2, output);
    push_varint(payload_len as u64, output);
    append_varint_field(output, 1, reference.identifier);
    if let Some(value) = reference.deprecated_type {
        append_varint_field(output, 2, value as i64 as u64);
    }
    if let Some(value) = reference.deprecated_is_external {
        append_varint_field(output, 3, u64::from(value));
    }
}

fn encoded_reference_field_len(field_number: u32, reference: ThemeReferenceSnapshot) -> usize {
    encoded_field_len(field_number, reference_payload_len(reference))
}

fn reference_payload_len(reference: ThemeReferenceSnapshot) -> usize {
    let mut length = encoded_varint_field_len(1, reference.identifier);
    if let Some(value) = reference.deprecated_type {
        length += encoded_varint_field_len(2, value as i64 as u64);
    }
    if let Some(value) = reference.deprecated_is_external {
        length += encoded_varint_field_len(3, u64::from(value));
    }
    length
}

fn encoded_varint_field_len(field_number: u32, value: u64) -> usize {
    varint_len(u64::from(field_number) << 3) + varint_len(value)
}

fn push_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn encoded_field_len(field_number: u32, payload_len: usize) -> usize {
    varint_len((u64::from(field_number) << 3) | 2) + varint_len(payload_len as u64) + payload_len
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn field_overflow(maximum: usize) -> DecodeError {
    DecodeError::resource(DecodeLimit::Fields, usize::MAX, maximum)
}

fn work_overflow(maximum: usize) -> DecodeError {
    DecodeError::resource(DecodeLimit::WorkBytes, usize::MAX, maximum)
}

fn output_overflow(maximum: usize) -> DecodeError {
    DecodeError::resource(DecodeLimit::OutputBytes, usize::MAX, maximum)
}

fn retained_overflow(maximum: usize) -> DecodeError {
    DecodeError::resource(DecodeLimit::RetainedBytes, usize::MAX, maximum)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{tsp, tss, tswp};
    use prost::Message as _;

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            deprecated_type: None,
            deprecated_is_external: None,
        }
    }

    fn field(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        append_raw_field(&mut output, number, payload);
        output
    }

    fn theme_source() -> Vec<u8> {
        let base = tss::ThemeArchive {
            legacy_stylesheet: Some(reference(11)),
            theme_identifier: Some("bounded-theme".to_owned()),
            document_stylesheet: Some(reference(12)),
            old_uuids_for_preset_replacements: vec![tsp::Uuid {
                lower: 101,
                upper: 102,
            }],
            new_uuids_for_preset_replacements: vec![tsp::Uuid {
                lower: 201,
                upper: 202,
            }],
            color_presets: vec![tsp::Color {
                model: 1,
                r: Some(1.0),
                g: Some(0.5),
                b: Some(0.0),
                rgbspace: Some(1),
                a: Some(1.0),
                c: None,
                m: None,
                y: None,
                k: None,
                w: None,
            }],
        }
        .encode_to_vec();
        let text = tswp::ThemePresetsArchive {
            list_style_presets: vec![reference(21)],
            ..Default::default()
        }
        .encode_to_vec();
        let mut inner = base;
        inner.extend(field(TEXT_PRESETS_FIELD, &text));
        inner.extend(field(APPLICATION_PRESETS_FIELD, b""));
        let mut inner_with_unknown = inner;
        inner_with_unknown.extend(field(333, b"base-unknown"));
        let mut wrapper = field(THEME_SUPER_FIELD, &inner_with_unknown);
        wrapper.extend(field(77, b"wrapper-unknown"));
        wrapper
    }

    #[test]
    fn differential_base_and_extension_projection_matches_prost() {
        let source = theme_source();
        let decoded = decode_theme_archive(&source, DecodeOptions::for_source(&source))
            .expect("theme decode");
        let inner = decoded.theme_raw();
        let prost_base = tss::ThemeArchive::decode(inner).expect("prost base decode");
        assert_eq!(
            decoded.theme_identifier(),
            prost_base.theme_identifier.as_deref()
        );
        let reference_payload = prost_base
            .legacy_stylesheet
            .as_ref()
            .map(|value| value.encode_to_vec());
        assert_eq!(decoded.legacy_stylesheet(), reference_payload.as_deref());
        assert_eq!(
            decoded.text_presets(),
            Some(field_payload(inner, TEXT_PRESETS_FIELD))
        );
        assert_eq!(
            decoded
                .extension_references(ThemeExtension::Text, 1)
                .collect::<Vec<_>>(),
            vec![ThemeReferenceSnapshot {
                identifier: 21,
                deprecated_type: None,
                deprecated_is_external: None,
            }]
        );
        assert_eq!(decoded.old_uuid_payloads().count(), 1);
        assert_eq!(decoded.new_uuid_payloads().count(), 1);
        assert_eq!(decoded.color_preset_payloads().count(), 1);
        assert_eq!(decoded.application_presets(), Some(&[][..]));
    }

    #[test]
    fn typed_reference_list_rewrite_preserves_complex_and_unknown_extension_fields() {
        let source = theme_source();
        let decoded = decode_theme_archive(&source, DecodeOptions::for_source(&source))
            .expect("theme decode");
        let mut text = decoded.text_presets().expect("text extension").to_vec();
        text.extend(field(77, b"future-text"));
        let references = [
            ThemeReferenceSnapshot::new(88),
            ThemeReferenceSnapshot {
                identifier: 89,
                deprecated_type: Some(-7),
                deprecated_is_external: Some(true),
            },
        ];
        let rewritten = rewrite_theme_extension_reference_list(
            &text,
            ThemeExtension::Text,
            1,
            &references,
            DecodeOptions::for_source(&text)
                .with_max_output_bytes(256)
                .with_max_scratch_bytes(256)
                .with_max_retained_bytes(512),
        )
        .expect("reference list rewrite");
        assert_eq!(
            ReferenceFields::new(&rewritten, 1).collect::<Vec<_>>(),
            references.to_vec()
        );
        assert_eq!(
            RawFields::new(&rewritten, RawFieldScope::Base).collect::<Vec<_>>(),
            vec![field(77, b"future-text").as_slice()]
        );
    }

    #[test]
    fn typed_reference_list_rewrite_preserves_unknown_nested_reference_fields() {
        let mut first_reference = ThemeReferenceSnapshot::new(88).encode().unwrap();
        first_reference.extend(field(77, b"removed-reference-future"));
        let mut retained_reference = ThemeReferenceSnapshot::new(89).encode().unwrap();
        retained_reference.extend(field(78, b"retained-reference-future"));
        let source = [
            field(1, &first_reference),
            field(1, &retained_reference),
            field(79, b"future-extension"),
        ]
        .concat();

        let rewritten = rewrite_theme_extension_reference_list(
            &source,
            ThemeExtension::Text,
            1,
            &[ThemeReferenceSnapshot::new(89)],
            DecodeOptions::for_source(&source)
                .with_max_output_bytes(512)
                .with_max_scratch_bytes(512)
                .with_max_retained_bytes(1024),
        )
        .expect("reference list rewrite");

        assert!(
            rewritten
                .windows(field(1, &retained_reference).len())
                .any(|window| window == field(1, &retained_reference).as_slice())
        );
        assert!(
            rewritten
                .windows(field(78, b"retained-reference-future").len())
                .any(|window| window == field(78, b"retained-reference-future").as_slice())
        );
        assert!(
            rewritten
                .windows(field(79, b"future-extension").len())
                .any(|window| window == field(79, b"future-extension").as_slice())
        );
        assert_eq!(
            ReferenceFields::new(&rewritten, 1).collect::<Vec<_>>(),
            vec![ThemeReferenceSnapshot::new(89)]
        );
    }

    #[test]
    fn typed_reference_list_rewrite_preserves_unknown_nested_fields_after_reorder() {
        let mut first_reference = ThemeReferenceSnapshot::new(88).encode().unwrap();
        first_reference.extend(field(77, b"first-reference-future"));
        let mut second_reference = ThemeReferenceSnapshot::new(89).encode().unwrap();
        second_reference.extend(field(78, b"second-reference-future"));
        let source = [field(1, &first_reference), field(1, &second_reference)].concat();

        let rewritten = rewrite_theme_extension_reference_list(
            &source,
            ThemeExtension::Text,
            1,
            &[
                ThemeReferenceSnapshot::new(89),
                ThemeReferenceSnapshot::new(88),
            ],
            DecodeOptions::for_source(&source)
                .with_max_output_bytes(512)
                .with_max_scratch_bytes(512)
                .with_max_retained_bytes(1024),
        )
        .expect("reference list reorder");

        assert_eq!(
            ReferenceFields::new(&rewritten, 1).collect::<Vec<_>>(),
            vec![
                ThemeReferenceSnapshot::new(89),
                ThemeReferenceSnapshot::new(88),
            ]
        );
        let expected = [field(1, &second_reference), field(1, &first_reference)].concat();
        assert_eq!(rewritten, expected);
    }

    #[test]
    fn typed_reference_list_rewrite_precharges_final_emission_walk() {
        let mut first_reference = ThemeReferenceSnapshot::new(88).encode().unwrap();
        first_reference.extend(field(77, b"first-reference-future"));
        let mut second_reference = ThemeReferenceSnapshot::new(89).encode().unwrap();
        second_reference.extend(field(78, b"second-reference-future"));
        let source = [field(1, &first_reference), field(1, &second_reference)].concat();
        let references = [
            ThemeReferenceSnapshot::new(89),
            ThemeReferenceSnapshot::new(88),
        ];
        let options = DecodeOptions::for_source(&source)
            .with_max_output_bytes(512)
            .with_max_scratch_bytes(512)
            .with_max_retained_bytes(1024)
            .with_max_work_bytes(1);
        assert_eq!(
            rewrite_theme_extension_reference_list(
                &source,
                ThemeExtension::Text,
                1,
                &references,
                options,
            )
            .expect_err("final extension emission walk must be admitted")
            .resource_limit(),
            Some(DecodeLimit::WorkBytes)
        );
    }

    #[test]
    fn unknown_fields_and_noop_rewrite_are_byte_exact() {
        let source = theme_source();
        let decoded = decode_theme_archive(&source, DecodeOptions::for_source(&source))
            .expect("theme decode");
        assert_eq!(
            decoded.unknown_wrapper_fields().collect::<Vec<_>>(),
            vec![field(77, b"wrapper-unknown").as_slice()]
        );
        assert_eq!(
            decoded.unknown_base_fields().collect::<Vec<_>>(),
            vec![field(333, b"base-unknown").as_slice()]
        );
        let output = rewrite_theme_archive(
            &source,
            ThemeArchiveWrite::new(),
            DecodeOptions::for_source(&source),
        )
        .expect("no-op rewrite");
        assert_eq!(output, source);
    }

    #[test]
    fn extension_rewrite_preserves_other_families_and_unknowns() {
        let source = theme_source();
        let replacement = tswp::ThemePresetsArchive {
            paragraph_style_presets: vec![reference(88)],
            ..Default::default()
        }
        .encode_to_vec();
        let write = ThemeArchiveWrite::new().with_text_presets(&replacement);
        let output = rewrite_theme_archive(&source, write, DecodeOptions::for_source(&source))
            .expect("extension rewrite");
        let decoded = decode_theme_archive(&output, DecodeOptions::for_source(&output))
            .expect("rewritten decode");
        assert_eq!(decoded.text_presets(), Some(replacement.as_slice()));
        assert_eq!(decoded.application_presets(), Some(&[][..]));
        assert_eq!(
            decoded.unknown_wrapper_fields().collect::<Vec<_>>(),
            vec![field(77, b"wrapper-unknown").as_slice()]
        );
        assert_eq!(
            decoded.unknown_base_fields().collect::<Vec<_>>(),
            vec![field(333, b"base-unknown").as_slice()]
        );
    }

    #[test]
    fn archive_reference_list_write_is_source_preserving_and_single_candidate() {
        let source = theme_source();
        let mut reference_payload = ThemeReferenceSnapshot::new(21).encode().unwrap();
        reference_payload.extend(field(88, b"future-reference"));
        let text = field(1, &reference_payload);
        let source = rewrite_theme_archive(
            &source,
            ThemeArchiveWrite::new().with_text_presets(&text),
            DecodeOptions::for_source(&source),
        )
        .expect("install source reference");

        let references = [ThemeReferenceSnapshot::new(21)];
        let unchanged = rewrite_theme_archive(
            &source,
            ThemeArchiveWrite::new().with_reference_list(ThemeExtension::Text, 1, &references),
            DecodeOptions::for_source(&source),
        )
        .expect("source-preserving reference write");
        assert_eq!(unchanged, source);
        assert!(
            unchanged
                .windows(field(88, b"future-reference").len())
                .any(|window| window == field(88, b"future-reference").as_slice())
        );

        let changed_references = [ThemeReferenceSnapshot::new(22)];
        let changed = rewrite_theme_archive(
            &source,
            ThemeArchiveWrite::new().with_reference_list(
                ThemeExtension::Text,
                1,
                &changed_references,
            ),
            DecodeOptions::for_source(&source),
        )
        .expect("changed reference write");
        let changed_snapshot =
            decode_theme_archive(&changed, DecodeOptions::for_source(&changed)).expect("readback");
        assert_eq!(
            changed_snapshot
                .extension_references(ThemeExtension::Text, 1)
                .collect::<Vec<_>>(),
            changed_references.to_vec()
        );
        assert!(
            !changed
                .windows(field(88, b"future-reference").len())
                .any(|window| window == field(88, b"future-reference").as_slice())
        );
    }

    #[test]
    fn archive_reference_list_exact_noop_reports_unchanged() {
        let source = theme_source();
        let references = [ThemeReferenceSnapshot::new(21)];
        let (snapshot, report) =
            decode_theme_archive_with_report(&source, DecodeOptions::for_source(&source))
                .expect("source comparison report");
        let (matches, comparison_report) = theme_reference_list_matches_with_report(
            snapshot,
            ThemeExtension::Text,
            1,
            &references,
            DecodeOptions::for_source(&source),
            report,
        )
        .expect("bounded source comparison");
        assert!(matches);
        assert!(comparison_report.work_bytes() > report.work_bytes());
        let (unchanged, report) = rewrite_theme_archive_with_report(
            &source,
            ThemeArchiveWrite::new().with_reference_list(ThemeExtension::Text, 1, &references),
            DecodeOptions::for_source(&source),
        )
        .expect("source-preserving reference report");
        assert_eq!(unchanged, source);
        assert!(!report.changed());

        let changed_references = [ThemeReferenceSnapshot::new(22)];
        let (_, changed_report) = rewrite_theme_archive_with_report(
            &source,
            ThemeArchiveWrite::new().with_reference_list(
                ThemeExtension::Text,
                1,
                &changed_references,
            ),
            DecodeOptions::for_source(&source),
        )
        .expect("changed reference report");
        assert!(changed_report.changed());
    }

    #[test]
    fn source_bound_reports_reject_same_length_foreign_sources() {
        let source = theme_source();
        let mut foreign = source.clone();
        let identifier_byte = foreign
            .iter()
            .position(|byte| *byte == b'b')
            .expect("theme fixture identifier");
        foreign[identifier_byte] = b'c';
        assert_eq!(foreign.len(), source.len());

        let (_, report) =
            decode_theme_archive_with_report(&source, DecodeOptions::for_source(&source))
                .expect("source report");
        let foreign_snapshot = decode_theme_archive(&foreign, DecodeOptions::for_source(&foreign))
            .expect("same-shaped foreign source");
        let error = admit_theme_archive_copy(&foreign, report, DecodeOptions::for_source(&foreign))
            .expect_err("copy admission must authenticate the source region");
        assert!(
            matches!(error.kind(), DecodeErrorKind::InvalidFormat(message) if message.contains("does not belong"))
        );

        let error = theme_reference_list_matches_with_report(
            foreign_snapshot,
            ThemeExtension::Text,
            1,
            &[ThemeReferenceSnapshot::new(21)],
            DecodeOptions::for_source(&foreign),
            report,
        )
        .expect_err("reference comparison must authenticate the source region");
        assert!(
            matches!(error.kind(), DecodeErrorKind::InvalidFormat(message) if message.contains("does not belong"))
        );

        let error = prepare_theme_archive_rewrite_with_prior_report(
            &foreign,
            ThemeArchiveWrite::new(),
            DecodeOptions::for_source(&foreign),
            Some(report),
        )
        .expect_err("rewrite preparation must authenticate the source region");
        assert!(
            matches!(error.kind(), DecodeErrorKind::InvalidFormat(message) if message.contains("does not belong"))
        );
    }

    #[test]
    fn archive_reference_reorder_duplicates_precharge_every_emission_walk() {
        let base = theme_source();
        let mut first = ThemeReferenceSnapshot::new(88).encode().unwrap();
        first.extend(field(77, b"first-reference-future"));
        let mut second = ThemeReferenceSnapshot::new(89).encode().unwrap();
        second.extend(field(78, b"second-reference-future"));
        let text = [field(1, &first), field(1, &second), field(1, &first)].concat();
        let source = rewrite_theme_archive(
            &base,
            ThemeArchiveWrite::new().with_text_presets(&text),
            DecodeOptions::for_source(&base),
        )
        .expect("install duplicate reference source");

        let references = [
            ThemeReferenceSnapshot::new(89),
            ThemeReferenceSnapshot::new(88),
            ThemeReferenceSnapshot::new(88),
        ];
        let write =
            ThemeArchiveWrite::new().with_reference_list(ThemeExtension::Text, 1, &references);
        let generous = DecodeOptions::for_source(&source)
            .with_max_fields(1_000_000)
            .with_max_work_bytes(1_000_000)
            .with_max_output_bytes(1_000_000)
            .with_max_retained_bytes(2_000_000)
            .with_max_scratch_bytes(1_000_000);
        let prepared = prepare_theme_archive_rewrite(&source, write, generous)
            .expect("prepare duplicate reorder");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("emit duplicate reorder")
            .into_bytes();
        let expected_text = [field(1, &second), field(1, &first), field(1, &first)].concat();
        let decoded = decode_theme_archive(&output, DecodeOptions::for_source(&output))
            .expect("read reordered duplicate output");
        assert_eq!(decoded.text_presets(), Some(expected_text.as_slice()));

        let tight = generous.with_max_work_bytes(requirements.work_bytes.saturating_sub(1));
        assert_eq!(
            prepare_theme_archive_rewrite(&source, write, tight)
                .expect_err("reorder matcher work must be precharged")
                .resource_limit(),
            Some(DecodeLimit::WorkBytes)
        );
    }

    #[test]
    fn admitted_copy_charges_source_report_and_owned_bytes() {
        let source = theme_source();
        let generous = DecodeOptions::for_source(&source);
        let (_, report) =
            decode_theme_archive_with_report(&source, generous).expect("source report");

        let work_limited = generous.with_max_work_bytes(report.work_bytes());
        assert_eq!(
            admit_theme_archive_copy(&source, report, work_limited)
                .expect_err("copy work must be aggregate")
                .resource_limit(),
            Some(DecodeLimit::WorkBytes)
        );

        let scratch_limited = generous.with_max_scratch_bytes(source.len().saturating_sub(1));
        assert_eq!(
            admit_theme_archive_copy(&source, report, scratch_limited)
                .expect_err("copy scratch must be admitted")
                .resource_limit(),
            Some(DecodeLimit::ScratchBytes)
        );

        let retained_limited = generous.with_max_retained_bytes(source.len());
        assert_eq!(
            admit_theme_archive_copy(&source, report, retained_limited)
                .expect_err("copy retained bytes must include source")
                .resource_limit(),
            Some(DecodeLimit::RetainedBytes)
        );

        assert_eq!(
            admit_theme_archive_copy(&source, report, generous).expect("admitted copy"),
            source
        );
    }

    #[test]
    fn reference_projection_uses_decode_and_projection_ledger() {
        let source = theme_source();
        let route = [ThemeExtension::Text];
        let generous = DecodeOptions::for_source(&source);
        let (_, report) =
            decode_theme_archive_with_report(&source, generous).expect("source report");
        let references =
            project_theme_references(&source, &route, |_| (ThemeExtension::Text, 1), generous)
                .expect("reference projection");
        assert_eq!(references, vec![ThemeReferenceSnapshot::new(21)]);

        let work_limited = generous.with_max_work_bytes(report.work_bytes());
        assert_eq!(
            project_theme_references(&source, &route, |_| (ThemeExtension::Text, 1), work_limited)
                .expect_err("projection traversal must share work budget")
                .resource_limit(),
            Some(DecodeLimit::WorkBytes)
        );

        let scratch_limited =
            generous.with_max_scratch_bytes(size_of::<ThemeReferenceSnapshot>().saturating_sub(1));
        assert_eq!(
            project_theme_references(
                &source,
                &route,
                |_| (ThemeExtension::Text, 1),
                scratch_limited,
            )
            .expect_err("projection vector must be scratch admitted")
            .resource_limit(),
            Some(DecodeLimit::ScratchBytes)
        );

        assert_eq!(
            first_theme_reference(&source, ThemeExtension::Text, 1, generous,)
                .expect("first reference projection"),
            Some(ThemeReferenceSnapshot::new(21))
        );
    }

    #[test]
    fn prior_decode_report_is_debited_before_reference_candidate_allocation() {
        let source = theme_source();
        let references = [ThemeReferenceSnapshot::new(22)];
        let write =
            ThemeArchiveWrite::new().with_reference_list(ThemeExtension::Text, 1, &references);
        let generous = DecodeOptions::for_source(&source)
            .with_max_fields(source.len().saturating_mul(64))
            .with_max_work_bytes(source.len().saturating_mul(256));
        let (_, prior_report) =
            decode_theme_archive_with_report(&source, generous).expect("source report");
        let prepared = prepare_theme_archive_rewrite_with_prior_report(
            &source,
            write,
            generous,
            Some(prior_report),
        )
        .expect("prepared rewrite");
        let requirements = prepared.execution_requirements();
        let error = prepare_theme_archive_rewrite_with_prior_report(
            &source,
            write,
            generous.with_max_fields(requirements.fields.saturating_sub(1)),
            Some(prior_report),
        )
        .expect_err("prior source pass must consume aggregate field budget");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::Fields));
    }

    #[test]
    fn malformed_duplicate_and_wrong_wire_inputs_fail_closed() {
        let duplicate_wrapper = [field(1, &[]), field(1, &[])].concat();
        let error = decode_theme_archive(
            &duplicate_wrapper,
            DecodeOptions::for_source(&duplicate_wrapper),
        )
        .expect_err("duplicate wrapper");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSS.ThemeArchive wrapper super")
        );

        let extension_a = field(DRAWING_PRESETS_FIELD, b"");
        let extension_b = field(DRAWING_PRESETS_FIELD, b"");
        let duplicate_inner = [extension_a, extension_b].concat();
        let duplicate_extension = field(THEME_SUPER_FIELD, &duplicate_inner);
        let error = decode_theme_archive(
            &duplicate_extension,
            DecodeOptions::for_source(&duplicate_extension),
        )
        .expect_err("duplicate extension");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSS.ThemeArchive extension field 100")
        );

        let wrong_wire = {
            let mut value = Vec::new();
            push_varint(0x08, &mut value);
            push_varint(1, &mut value);
            value
        };
        let error = decode_theme_archive(&wrong_wire, DecodeOptions::for_source(&wrong_wire))
            .expect_err("wrong wrapper wire");
        assert!(matches!(error.kind(), DecodeErrorKind::InvalidFormat(_)));

        let malformed_reference = {
            let extension = field(1, &[]);
            let inner = field(TEXT_PRESETS_FIELD, &extension);
            field(THEME_SUPER_FIELD, &inner)
        };
        let error = decode_theme_archive(
            &malformed_reference,
            DecodeOptions::for_source(&malformed_reference),
        )
        .expect_err("missing preset reference identifier");
        assert_eq!(
            error.missing_required_field(),
            Some("TSP.Reference.identifier")
        );

        let duplicate_required_reference_payload = [0x08, 0x01, 0x08, 0x02];
        let duplicate_required_reference = {
            let extension = field(1, &duplicate_required_reference_payload);
            let inner = field(TEXT_PRESETS_FIELD, &extension);
            field(THEME_SUPER_FIELD, &inner)
        };
        let error = decode_theme_archive(
            &duplicate_required_reference,
            DecodeOptions::for_source(&duplicate_required_reference),
        )
        .expect_err("duplicate required preset reference field");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSP.Reference.identifier")
        );

        let duplicate_optional_reference_payload = [0x08, 0x01, 0x10, 0x02, 0x10, 0x03];
        let duplicate_optional_reference = {
            let extension = field(1, &duplicate_optional_reference_payload);
            let inner = field(TEXT_PRESETS_FIELD, &extension);
            field(THEME_SUPER_FIELD, &inner)
        };
        let error = decode_theme_archive(
            &duplicate_optional_reference,
            DecodeOptions::for_source(&duplicate_optional_reference),
        )
        .expect_err("duplicate optional preset reference field");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSP.Reference.deprecated_type")
        );

        // A non-canonical tag/length varint must not be silently rewritten
        // into a canonical no-op candidate.  The strict pass rejects it
        // before either the Buffa projection or rewrite emitter runs.
        let noncanonical_wrapper = [0x8a, 0x00, 0x84, 0x00, 0x0a, 0x02, 0x08, 0x01];
        let error = decode_theme_archive(
            &noncanonical_wrapper,
            DecodeOptions::for_source(&noncanonical_wrapper),
        )
        .expect_err("non-canonical wrapper varints");
        assert_eq!(
            error.noncanonical_reason(),
            Some("protobuf varint is not canonically encoded")
        );
        let error = rewrite_theme_archive(
            &noncanonical_wrapper,
            ThemeArchiveWrite::new(),
            DecodeOptions::for_source(&noncanonical_wrapper),
        )
        .expect_err("non-canonical no-op rewrite");
        assert_eq!(
            error.noncanonical_reason(),
            Some("protobuf varint is not canonically encoded")
        );
    }

    #[test]
    fn reference_ids_and_known_scalars_are_strictly_admitted() {
        let reference = field(TEXT_PRESETS_FIELD, &field(1, &[0x08, 0x00]));
        let zero_id = field(THEME_SUPER_FIELD, &reference);
        let error = decode_theme_archive(&zero_id, DecodeOptions::for_source(&zero_id))
            .expect_err("zero native identifiers must be rejected");
        assert!(
            matches!(error.kind(), DecodeErrorKind::InvalidFormat(message) if message.contains("non-zero"))
        );

        let zero_rewrite = [ThemeReferenceSnapshot::new(0)];
        let source = theme_source();
        let error = prepare_theme_archive_rewrite(
            &source,
            ThemeArchiveWrite::new().with_reference_list(ThemeExtension::Text, 1, &zero_rewrite),
            DecodeOptions::for_source(&source),
        )
        .expect_err("zero native identifiers must be rejected before rewrite allocation");
        assert!(
            matches!(error.kind(), DecodeErrorKind::InvalidFormat(message) if message.contains("non-zero"))
        );

        let bool_two_payload = [0x08, 0x01, 0x18, 0x02];
        let bool_two = field(
            THEME_SUPER_FIELD,
            &field(TEXT_PRESETS_FIELD, &field(1, &bool_two_payload)),
        );
        let error = decode_theme_archive(&bool_two, DecodeOptions::for_source(&bool_two))
            .expect_err("boolean values other than zero or one must be rejected");
        assert_eq!(
            error.noncanonical_reason(),
            Some("TSP.Reference.deprecated_is_external bool scalar is not zero or one")
        );

        let int32_noncanonical = [0x08, 0x01, 0x10, 0x80, 0x80, 0x80, 0x80, 0x08];
        let int32_source = field(
            THEME_SUPER_FIELD,
            &field(TEXT_PRESETS_FIELD, &field(1, &int32_noncanonical)),
        );
        let error = decode_theme_archive(&int32_source, DecodeOptions::for_source(&int32_source))
            .expect_err("int32 values must use sign-extended canonical encoding");
        assert_eq!(
            error.noncanonical_reason(),
            Some("TSP.Reference.deprecated_type is not a sign-extended int32")
        );

        let color_noncanonical = [0x08, 0x80, 0x80, 0x80, 0x80, 0x08];
        let color_source = field(
            THEME_SUPER_FIELD,
            &field(COLOR_PRESETS_FIELD, &color_noncanonical),
        );
        let error = decode_theme_archive(&color_source, DecodeOptions::for_source(&color_source))
            .expect_err("color enum scalars must use canonical int32 encoding");
        assert_eq!(
            error.noncanonical_reason(),
            Some("TSP.Color.model is not a sign-extended int32")
        );
    }

    #[test]
    fn finite_limits_precharge_decode_and_rewrite() {
        let source = theme_source();
        let options = DecodeOptions::for_source(&source).with_max_message_bytes(source.len() - 1);
        let error = decode_theme_archive(&source, options).expect_err("input limit");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::InputBytes));

        let options = DecodeOptions::for_source(&source).with_max_fields(1);
        let error = decode_theme_archive(&source, options).expect_err("field limit");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::Fields));

        let replacement = vec![0x0a, 0x01, 0x08, 0x01];
        let write = ThemeArchiveWrite::new().with_text_presets(&replacement);
        let options = DecodeOptions::for_source(&source).with_max_output_bytes(1);
        let error = prepare_theme_archive_rewrite(&source, write, options)
            .expect_err("rewrite output limit");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::OutputBytes));

        let prepared = prepare_theme_archive_rewrite(
            &source,
            ThemeArchiveWrite::new(),
            DecodeOptions::for_source(&source),
        )
        .expect("prepare exact rewrite");
        let requirements = prepared.execution_requirements();
        let error = prepared
            .execute(
                RewriteExecutionLimits::exact(requirements)
                    .with_fields(requirements.fields.saturating_sub(1)),
            )
            .expect_err("field precharge");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::Fields));

        let prepared = prepare_theme_archive_rewrite(
            &source,
            ThemeArchiveWrite::new(),
            DecodeOptions::for_source(&source),
        )
        .expect("prepare exact work rewrite");
        let requirements = prepared.execution_requirements();
        let error = prepared
            .execute(
                RewriteExecutionLimits::exact(requirements)
                    .with_work_bytes(requirements.work_bytes.saturating_sub(1)),
            )
            .expect_err("work precharge");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::WorkBytes));
    }

    #[test]
    fn builder_validates_repeated_records_and_unknown_fields_before_allocation() {
        let legacy = reference(31).encode_to_vec();
        let document = reference(32).encode_to_vec();
        let old_uuid = tsp::Uuid {
            lower: 41,
            upper: 42,
        }
        .encode_to_vec();
        let new_uuid = tsp::Uuid {
            lower: 51,
            upper: 52,
        }
        .encode_to_vec();
        let color = tsp::Color {
            model: 1,
            r: Some(1.0),
            g: Some(0.25),
            b: Some(0.5),
            rgbspace: Some(1),
            a: None,
            c: None,
            m: None,
            y: None,
            k: None,
            w: None,
        }
        .encode_to_vec();
        let text = tswp::ThemePresetsArchive {
            list_style_presets: vec![reference(61)],
            ..Default::default()
        }
        .encode_to_vec();
        let base_unknown = field(333, b"base-builder-unknown");
        let wrapper_unknown = field(777, b"wrapper-builder-unknown");
        let old_uuids = [old_uuid.as_slice()];
        let new_uuids = [new_uuid.as_slice()];
        let colors = [color.as_slice()];
        let unknown_base = [base_unknown.as_slice()];
        let unknown_wrapper = [wrapper_unknown.as_slice()];
        let options = DecodeOptions::new(4096, 4096, 65_536, 16)
            .with_max_output_bytes(4096)
            .with_max_retained_bytes(8192)
            .with_max_scratch_bytes(4096);
        let output = ThemeArchiveBuilder::new()
            .with_legacy_stylesheet(&legacy)
            .with_theme_identifier("fresh-theme")
            .with_document_stylesheet(&document)
            .with_old_uuid_payloads(&old_uuids)
            .with_new_uuid_payloads(&new_uuids)
            .with_color_preset_payloads(&colors)
            .with_text_presets(&text)
            .with_unknown_base_fields(&unknown_base)
            .with_unknown_wrapper_fields(&unknown_wrapper)
            .encode(options)
            .expect("fresh theme builder");
        let decoded = decode_theme_archive(&output, DecodeOptions::for_source(&output))
            .expect("decode fresh theme");
        assert_eq!(decoded.theme_identifier(), Some("fresh-theme"));
        assert_eq!(decoded.legacy_stylesheet(), Some(legacy.as_slice()));
        assert_eq!(decoded.document_stylesheet(), Some(document.as_slice()));
        assert_eq!(decoded.old_uuid_payloads().count(), 1);
        assert_eq!(decoded.new_uuid_payloads().count(), 1);
        assert_eq!(decoded.color_preset_payloads().count(), 1);
        assert_eq!(decoded.text_presets(), Some(text.as_slice()));
        assert_eq!(
            decoded.unknown_base_fields().collect::<Vec<_>>(),
            vec![base_unknown.as_slice()]
        );
        assert_eq!(
            decoded.unknown_wrapper_fields().collect::<Vec<_>>(),
            vec![wrapper_unknown.as_slice()]
        );

        let error = ThemeArchiveBuilder::new()
            .with_legacy_stylesheet(&[0x08, 0x01, 0x08, 0x02])
            .encode(options)
            .expect_err("duplicate builder reference");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSP.Reference.identifier")
        );
        let error = ThemeArchiveBuilder::new()
            .with_theme_identifier("too-small-for-output-policy")
            .encode(options.with_max_output_bytes(1))
            .expect_err("builder output precharge");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::OutputBytes));
    }

    #[test]
    fn rewrite_readback_keeps_explicit_deep_unknown_group_budget() {
        let mut nested = Vec::new();
        for _ in 0..12 {
            push_varint((333u64 << 3) | 3, &mut nested);
        }
        for _ in 0..12 {
            push_varint((333u64 << 3) | 4, &mut nested);
        }
        // Keep groups directly in the base message. A length-delimited
        // unknown field would keep their bytes opaque and never test depth.
        let source = field(THEME_SUPER_FIELD, &nested);
        let options = DecodeOptions::new(4096, 4096, 65_536, 16)
            .with_max_output_bytes(4096)
            .with_max_retained_bytes(8192)
            .with_max_scratch_bytes(4096);
        let decoded = decode_theme_archive(&source, options).expect("deep theme decode");
        assert_eq!(decoded.unknown_base_fields().count(), 1);
        let output = rewrite_theme_archive(&source, ThemeArchiveWrite::new(), options)
            .expect("deep no-op rewrite");
        assert_eq!(output, source);
    }

    #[test]
    fn builder_rejects_deep_unknown_groups_before_reaching_malformed_tail() {
        let mut nested = Vec::new();
        for _ in 0..512 {
            push_varint((333u64 << 3) | 3, &mut nested);
        }
        // Deliberately unterminated: an unbounded preliminary traversal
        // would report malformed framing before consulting the depth budget.
        let unknown = [nested.as_slice()];
        let options = DecodeOptions::new(4096, 4096, 65_536, 4)
            .with_max_output_bytes(4096)
            .with_max_retained_bytes(8192)
            .with_max_scratch_bytes(4096);
        let error = ThemeArchiveBuilder::new()
            .with_unknown_base_fields(&unknown)
            .encode(options)
            .expect_err("depth admission precedes traversal of the malformed tail");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::Nesting));
        let error = ThemeArchiveBuilder::new()
            .with_unknown_wrapper_fields(&unknown)
            .encode(options.with_max_fields(1))
            .expect_err("field admission precedes traversal of the malformed tail");
        assert_eq!(error.resource_limit(), Some(DecodeLimit::Fields));
    }

    fn field_payload(source: &[u8], number: u32) -> &[u8] {
        let mut cursor = 0usize;
        while cursor < source.len() {
            let field = next_field(source, &mut cursor).expect("valid test field");
            if field.number == number {
                return &source[field.span.payload_start..field.span.payload_end];
            }
        }
        panic!("field {number} missing");
    }

    #[test]
    fn zero_reference_identifier_is_a_typed_error() {
        assert!(ThemeReferenceSnapshot::new(0).encode().is_err());
        assert!(!ThemeReferenceSnapshot::new(7).encode().unwrap().is_empty());
    }
}
