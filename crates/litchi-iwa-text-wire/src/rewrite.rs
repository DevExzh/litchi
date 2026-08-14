//! Bounded, byte-preserving mutation of native TSWP text storage.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Wire validation types stay adjacent to the passes that consume them"
)]
#![allow(
    clippy::struct_field_names,
    reason = "Every retained maximum has a matching explicit accessor"
)]

use std::{mem::size_of, ops::Range};

#[cfg(test)]
std::thread_local! {
    static STORAGE_EXECUTION_ENTRIES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

use litchi_iwa_common::varint::{
    decode_varint_from_bytes, encode_varint_into, encoded_len as varint_len,
};
use litchi_iwa_protos::text_storage_codec;

const ROOT_TEXT_FIELD: u32 = 3;
const TABLE_ENTRY_FIELD: u32 = 1;
const MAX_KNOWN_ROOT_FIELD: usize = 28;
const MAX_SCHEMA_NESTING: usize = 4;

/// Finite resource policy retained for the complete storage rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteLimits {
    max_message_bytes: usize,
    max_fields: usize,
    max_nesting: usize,
    max_fragments: usize,
    max_text_bytes: usize,
    max_table_entries: usize,
    max_object_references: usize,
    max_output_bytes: usize,
    max_rewrite_work: usize,
}

impl RewriteLimits {
    /// Absolute input-message ceiling.
    pub const MAX_MESSAGE_BYTES: usize = 512 * 1024 * 1024;
    /// Absolute aggregate parsed-field ceiling.
    pub const MAX_FIELDS: usize = 1_000_000;
    /// Absolute nested-message ceiling.
    pub const MAX_NESTING: usize = 64;
    /// Absolute field-3 occurrence ceiling.
    pub const MAX_FRAGMENTS: usize = 1_000_000;
    /// Absolute source/result UTF-8 text ceiling.
    pub const MAX_TEXT_BYTES: usize = 512 * 1024 * 1024;
    /// Absolute aggregate table-entry ceiling.
    pub const MAX_TABLE_ENTRIES: usize = 1_000_000;
    /// Absolute schema-declared reference-occurrence ceiling.
    pub const MAX_OBJECT_REFERENCES: usize = 1_000_000;
    /// Absolute encoded-output ceiling.
    pub const MAX_OUTPUT_BYTES: usize = 512 * 1024 * 1024;
    /// Absolute aggregate scanning/copying work ceiling.
    pub const MAX_REWRITE_WORK: usize = 2_147_483_647;

    /// Construct an explicit, non-bypassable rewrite profile.
    ///
    /// # Errors
    ///
    /// Returns [`RewriteError::InvalidLimit`] when a value is zero or exceeds
    /// its hard ceiling. A nesting limit below the four schema levels needed
    /// by storage, table, entry, and reference/range messages is also rejected.
    #[allow(
        clippy::too_many_arguments,
        reason = "Every independently retained resource bound is explicit at the trust boundary"
    )]
    pub fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_nesting: usize,
        max_fragments: usize,
        max_text_bytes: usize,
        max_table_entries: usize,
        max_object_references: usize,
        max_output_bytes: usize,
        max_rewrite_work: usize,
    ) -> RewriteResult<Self> {
        let nesting = checked_limit("nesting", max_nesting, Self::MAX_NESTING)?;
        if nesting < MAX_SCHEMA_NESTING {
            return Err(RewriteError::InvalidLimit {
                field: "nesting",
                value: nesting,
                maximum: Self::MAX_NESTING,
            });
        }
        Ok(Self {
            max_message_bytes: checked_limit(
                "message bytes",
                max_message_bytes,
                Self::MAX_MESSAGE_BYTES,
            )?,
            max_fields: checked_limit("fields", max_fields, Self::MAX_FIELDS)?,
            max_nesting: nesting,
            max_fragments: checked_limit("text fragments", max_fragments, Self::MAX_FRAGMENTS)?,
            max_text_bytes: checked_limit("text bytes", max_text_bytes, Self::MAX_TEXT_BYTES)?,
            max_table_entries: checked_limit(
                "table entries",
                max_table_entries,
                Self::MAX_TABLE_ENTRIES,
            )?,
            max_object_references: checked_limit(
                "object references",
                max_object_references,
                Self::MAX_OBJECT_REFERENCES,
            )?,
            max_output_bytes: checked_limit(
                "output bytes",
                max_output_bytes,
                Self::MAX_OUTPUT_BYTES,
            )?,
            max_rewrite_work: checked_limit(
                "rewrite work",
                max_rewrite_work,
                Self::MAX_REWRITE_WORK,
            )?,
        })
    }

    /// Maximum encoded input bytes.
    #[must_use]
    pub const fn max_message_bytes(self) -> usize {
        self.max_message_bytes
    }

    /// Maximum aggregate fields across schema-directed nested messages.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Maximum nested-message depth.
    #[must_use]
    pub const fn max_nesting(self) -> usize {
        self.max_nesting
    }

    /// Maximum field-3 fragments.
    #[must_use]
    pub const fn max_fragments(self) -> usize {
        self.max_fragments
    }

    /// Maximum aggregate source and resulting UTF-8 text bytes.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Maximum entries across every positional table.
    #[must_use]
    pub const fn max_table_entries(self) -> usize {
        self.max_table_entries
    }

    /// Maximum schema-declared reference occurrences.
    #[must_use]
    pub const fn max_object_references(self) -> usize {
        self.max_object_references
    }

    /// Maximum encoded rewritten bytes.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Maximum aggregate bytes scanned or copied by the transaction.
    #[must_use]
    pub const fn max_rewrite_work(self) -> usize {
        self.max_rewrite_work
    }
}

impl Default for RewriteLimits {
    fn default() -> Self {
        Self {
            max_message_bytes: 64 * 1024 * 1024,
            max_fields: 100_000,
            max_nesting: 8,
            max_fragments: 100_000,
            max_text_bytes: 64 * 1024 * 1024,
            max_table_entries: 100_000,
            max_object_references: 100_000,
            max_output_bytes: 64 * 1024 * 1024,
            max_rewrite_work: 512 * 1024 * 1024,
        }
    }
}

/// Attribute-table behavior when replacement text equals the selected text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum RewriteBehavior {
    /// Treat equal selected/replacement text as an exact semantic no-op.
    PreserveOnEqualText,
    /// Apply native selection-replacement table semantics even when the text
    /// bytes are equal. This retains legacy editor behavior.
    ReplaceSelection,
}

/// Resource and text facts proven by full storage validation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StorageValidation {
    utf8_len: usize,
    utf16_len: usize,
    fragments: usize,
    fields: usize,
    table_entries: usize,
    reference_occurrences: usize,
    validation_work: usize,
    has_unknown_wire_fields: bool,
}

impl StorageValidation {
    /// Aggregate source text length in UTF-8 bytes.
    #[must_use]
    pub const fn utf8_len(self) -> usize {
        self.utf8_len
    }

    /// Aggregate source text length in UTF-16 code units.
    #[must_use]
    pub const fn utf16_len(self) -> usize {
        self.utf16_len
    }

    /// Number of repeated field-3 text fragments.
    #[must_use]
    pub const fn fragments(self) -> usize {
        self.fragments
    }

    /// Aggregate known and unknown fields across traversed messages.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate entries across positional tables.
    #[must_use]
    pub const fn table_entries(self) -> usize {
        self.table_entries
    }

    /// Aggregate schema-declared reference occurrences, including duplicates.
    #[must_use]
    pub const fn reference_occurrences(self) -> usize {
        self.reference_occurrences
    }

    /// Conservative aggregate work charged by strict validation.
    #[must_use]
    pub const fn validation_work(self) -> usize {
        self.validation_work
    }

    /// Whether any traversed message contained an unrecognized field.
    #[must_use]
    pub const fn has_unknown_wire_fields(self) -> bool {
        self.has_unknown_wire_fields
    }
}

/// A successfully rewritten storage payload and its exact reference delta.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StorageRewrite {
    bytes: Vec<u8>,
    before_utf16_len: usize,
    after_utf16_len: usize,
    object_references_before: Vec<u64>,
    object_references_after: Vec<u64>,
    object_reference_occurrences_before: Vec<u64>,
    object_reference_occurrences_after: Vec<u64>,
    removed_object_references: Vec<u64>,
    removed_object_references_by_field: Vec<RemovedObjectReference>,
    has_unknown_wire_fields: bool,
    changed: bool,
    execution_report: StorageRewriteExecutionReport,
}

/// Candidate-production requirements known after strict, output-free
/// planning. Every retained and temporary vector is bounded by the exact
/// encoded output length or the exact source reference-occurrence count.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StorageRewriteExecutionRequirements {
    output_bytes: usize,
    retained_elements: usize,
    retained_bytes: usize,
    peak_scratch_bytes: usize,
    allocations: usize,
    work: usize,
    reference_occurrences: usize,
}

impl StorageRewriteExecutionRequirements {
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub const fn retained_elements(self) -> usize {
        self.retained_elements
    }

    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub const fn peak_scratch_bytes(self) -> usize {
        self.peak_scratch_bytes
    }

    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    #[must_use]
    pub const fn work(self) -> usize {
        self.work
    }

    #[must_use]
    pub const fn reference_occurrences(self) -> usize {
        self.reference_occurrences
    }
}

/// Independent execution ceilings checked before the first candidate or
/// reference vector is allocated.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StorageRewriteExecutionLimits {
    pub max_output_bytes: usize,
    pub max_retained_elements: usize,
    pub max_retained_bytes: usize,
    pub max_peak_scratch_bytes: usize,
    pub max_allocations: usize,
    pub max_work: usize,
}

/// Exact post-execution resource report. Retained axes use actual vector
/// capacities; peak scratch includes the simultaneously live intermediate
/// reference projections.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct StorageRewriteExecutionReport {
    pub retained_elements: usize,
    pub retained_bytes: usize,
    pub peak_scratch_bytes: usize,
    pub allocations: usize,
    pub work: usize,
}

/// Strict output-free rewrite plan. It borrows source and replacement text;
/// execution alone materializes candidate bytes or reference vectors.
pub struct PreparedStorageRewrite<'source, 'replacement> {
    source: &'source [u8],
    range: Range<usize>,
    replacement: &'replacement str,
    text: TextPlan,
    root: RootPlan,
    limits: RewriteLimits,
    requirements: StorageRewriteExecutionRequirements,
}

impl PreparedStorageRewrite<'_, '_> {
    #[must_use]
    pub const fn execution_requirements(&self) -> StorageRewriteExecutionRequirements {
        self.requirements
    }

    #[must_use]
    pub const fn changed(&self) -> bool {
        self.root.expected_changed
    }

    /// Execute the already-validated rewrite under independent finite limits.
    ///
    /// # Errors
    ///
    /// Refuses any insufficient axis before allocating the candidate buffer.
    pub fn execute(self, limits: StorageRewriteExecutionLimits) -> RewriteResult<StorageRewrite> {
        preflight_storage_execution(self.requirements, limits)?;
        #[cfg(test)]
        STORAGE_EXECUTION_ENTRIES.with(|entries| entries.set(entries.get() + 1));
        execute_prepared_storage_rewrite(self)
    }
}

/// One schema-declared reference removed from a specific storage field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct RemovedObjectReference {
    storage_field_number: u32,
    identifier: u64,
}

impl RemovedObjectReference {
    /// The root `TSWP.StorageArchive` field that owned the reference.
    #[must_use]
    pub const fn storage_field_number(self) -> u32 {
        self.storage_field_number
    }

    /// Referenced IWA object identifier.
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }
}

impl StorageRewrite {
    /// Borrow the rewritten storage bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Consume the result and return the rewritten storage bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    /// Whether any source byte changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    #[must_use]
    pub const fn execution_report(&self) -> StorageRewriteExecutionReport {
        self.execution_report
    }

    /// Aggregate source text length in UTF-16 code units.
    #[must_use]
    pub const fn before_utf16_len(&self) -> usize {
        self.before_utf16_len
    }

    /// Aggregate rewritten text length in UTF-16 code units.
    #[must_use]
    pub const fn after_utf16_len(&self) -> usize {
        self.after_utf16_len
    }

    /// Sorted, deduplicated schema-declared references before the rewrite.
    #[must_use]
    pub fn object_references_before(&self) -> &[u64] {
        &self.object_references_before
    }

    /// Sorted, deduplicated schema-declared references after the rewrite.
    #[must_use]
    pub fn object_references_after(&self) -> &[u64] {
        &self.object_references_after
    }

    /// Sorted schema-declared reference identifiers before the rewrite,
    /// retaining duplicate occurrences for metadata parity checks.
    #[must_use]
    pub fn object_reference_occurrences_before(&self) -> &[u64] {
        &self.object_reference_occurrences_before
    }

    /// Sorted schema-declared reference identifiers after the rewrite,
    /// retaining duplicate occurrences for metadata parity checks.
    #[must_use]
    pub fn object_reference_occurrences_after(&self) -> &[u64] {
        &self.object_reference_occurrences_after
    }

    /// Sorted, deduplicated references present before but absent afterward.
    #[must_use]
    pub fn removed_object_references(&self) -> &[u64] {
        &self.removed_object_references
    }

    /// Sorted, deduplicated `(storage field, identifier)` removals.
    #[must_use]
    pub fn removed_object_references_by_field(&self) -> &[RemovedObjectReference] {
        &self.removed_object_references_by_field
    }

    /// Whether any traversed source message contained an unrecognized field.
    #[must_use]
    pub const fn has_unknown_wire_fields(&self) -> bool {
        self.has_unknown_wire_fields
    }
}

/// Why a raw `TSWP.StorageArchive` rewrite was refused.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum RewriteError {
    /// A retained limit was zero, too small for the schema, or above its hard ceiling.
    #[error(
        "invalid text rewrite limit {field}: {value}, expected a supported value up to {maximum}"
    )]
    InvalidLimit {
        /// Limit field.
        field: &'static str,
        /// Supplied value.
        value: usize,
        /// Hard maximum.
        maximum: usize,
    },
    /// A finite resource bound was exceeded before output allocation.
    #[error("text rewrite {resource} limit exceeded: observed {observed}, limit {limit}")]
    LimitExceeded {
        /// Limited resource.
        resource: &'static str,
        /// Observed or required amount.
        observed: usize,
        /// Configured maximum.
        limit: usize,
    },
    /// The requested UTF-16 range starts after it ends.
    #[error("text replacement range {start}..{end} is reversed")]
    ReversedRange {
        /// Requested start.
        start: usize,
        /// Requested end.
        end: usize,
    },
    /// A requested UTF-16 endpoint exceeds the source text.
    #[error("text replacement UTF-16 index {index} exceeds source length {length}")]
    RangeOutOfBounds {
        /// Requested endpoint.
        index: usize,
        /// Source UTF-16 length.
        length: usize,
    },
    /// A requested UTF-16 endpoint falls between one scalar's surrogate pair.
    #[error("text replacement UTF-16 index {index} splits a surrogate pair")]
    SurrogateSplit {
        /// Requested endpoint.
        index: usize,
    },
    /// Checked text or index arithmetic overflowed.
    #[error("text rewrite arithmetic overflow while computing {context}")]
    ArithmeticOverflow {
        /// Failed calculation.
        context: &'static str,
    },
    /// Recognized wire data violated its strict schema contract.
    #[error("invalid TSWP storage wire data: {0}")]
    InvalidFormat(String),
    /// Buffa rejected a source that passed raw structural preflight.
    #[error("TSWP storage Buffa projection failed: {0}")]
    Projection(String),
    /// A fallible allocation failed.
    #[error("could not allocate {amount} elements for {resource}")]
    Allocation {
        /// Allocation purpose.
        resource: &'static str,
        /// Requested elements or bytes.
        amount: usize,
    },
}

/// Result type for bounded storage rewrites.
pub type RewriteResult<T> = Result<T, RewriteError>;

#[derive(Debug, Clone, Copy)]
struct RawField<'source> {
    source: &'source [u8],
    number: u32,
    wire_type: u8,
    start: usize,
    key_end: usize,
    payload_start: usize,
    end: usize,
}

impl<'source> RawField<'source> {
    fn canonical_payload(
        self,
        expected_wire: u8,
        label: &'static str,
    ) -> RewriteResult<&'source [u8]> {
        if self.wire_type != expected_wire {
            return Err(RewriteError::InvalidFormat(format!(
                "{label} field {} has wire type {}; expected {expected_wire}",
                self.number, self.wire_type
            )));
        }
        Ok(self.payload())
    }

    fn canonical_varint(self, label: &'static str) -> RewriteResult<u64> {
        let payload = self.canonical_payload(0, label)?;
        let (value, length) = decode_varint_from_bytes(payload).map_err(|error| {
            RewriteError::InvalidFormat(format!(
                "{label} field {} contains an invalid varint: {error}",
                self.number
            ))
        })?;
        if length != payload.len() {
            return Err(RewriteError::InvalidFormat(format!(
                "{label} field {} contains trailing varint bytes",
                self.number
            )));
        }
        Ok(value)
    }

    fn payload(self) -> &'source [u8] {
        &self.source[self.payload_start..self.end]
    }

    fn raw(self) -> &'source [u8] {
        &self.source[self.start..self.end]
    }

    fn key(self) -> &'source [u8] {
        &self.source[self.start..self.key_end]
    }
}

struct RawFields<'source> {
    source: &'source [u8],
    offset: usize,
}

impl<'source> RawFields<'source> {
    const fn new(source: &'source [u8]) -> Self {
        Self { source, offset: 0 }
    }

    fn next(&mut self) -> RewriteResult<Option<RawField<'source>>> {
        if self.offset == self.source.len() {
            return Ok(None);
        }
        let start = self.offset;
        let (key_value, key_length) =
            decode_varint_from_bytes(&self.source[start..]).map_err(|error| {
                RewriteError::InvalidFormat(format!("invalid protobuf key: {error}"))
            })?;
        let key_end = checked_add(start, key_length, "protobuf key offset")?;
        let number = u32::try_from(key_value >> 3)
            .map_err(|_error| invalid("protobuf field number exceeds u32"))?;
        if number == 0 || number > 0x1fff_ffff {
            return Err(invalid("protobuf field number is outside the valid range"));
        }
        let wire_type = u8::try_from(key_value & 7)
            .map_err(|_error| invalid("protobuf wire type exceeds u8"))?;
        let (payload_start, end) = match wire_type {
            0 => {
                let (_, length) =
                    decode_varint_from_bytes(&self.source[key_end..]).map_err(|error| {
                        invalid_owned(format!("invalid protobuf varint value: {error}"))
                    })?;
                (
                    key_end,
                    checked_add(key_end, length, "protobuf varint offset")?,
                )
            },
            1 => (key_end, checked_add(key_end, 8, "protobuf fixed64 offset")?),
            2 => {
                let (encoded_length, prefix_length) =
                    decode_varint_from_bytes(&self.source[key_end..]).map_err(|error| {
                        invalid_owned(format!("invalid protobuf length: {error}"))
                    })?;
                let payload_start = checked_add(key_end, prefix_length, "protobuf length prefix")?;
                let payload_length = usize::try_from(encoded_length)
                    .map_err(|_error| invalid("protobuf length exceeds usize"))?;
                (
                    payload_start,
                    checked_add(payload_start, payload_length, "protobuf payload range")?,
                )
            },
            5 => (key_end, checked_add(key_end, 4, "protobuf fixed32 offset")?),
            3 | 4 => return Err(invalid("deprecated protobuf groups are unsupported")),
            _ => return Err(invalid("invalid protobuf wire type")),
        };
        if end > self.source.len() {
            return Err(invalid("truncated protobuf field"));
        }
        self.offset = end;
        Ok(Some(RawField {
            source: self.source,
            number,
            wire_type,
            start,
            key_end,
            payload_start,
            end,
        }))
    }
}

#[derive(Debug, Clone, Copy)]
struct TextPreflight {
    fragments: usize,
    text_bytes: usize,
    utf16_len: usize,
}

#[derive(Debug, Clone, Copy)]
struct TextPlan {
    before_utf16_len: usize,
    after_utf16_len: usize,
    start_byte: usize,
    end_byte: usize,
    replacement_units: usize,
    result_bytes: usize,
    anchor_fragment: Option<usize>,
    no_op: bool,
}

fn checked_limit(field: &'static str, value: usize, maximum: usize) -> RewriteResult<usize> {
    if value == 0 || value > maximum {
        return Err(RewriteError::InvalidLimit {
            field,
            value,
            maximum,
        });
    }
    Ok(value)
}

fn invalid(message: &'static str) -> RewriteError {
    RewriteError::InvalidFormat(message.to_owned())
}

fn invalid_owned(message: String) -> RewriteError {
    RewriteError::InvalidFormat(message)
}

fn arithmetic(context: &'static str) -> RewriteError {
    RewriteError::ArithmeticOverflow { context }
}

fn checked_add(left: usize, right: usize, context: &'static str) -> RewriteResult<usize> {
    left.checked_add(right).ok_or_else(|| arithmetic(context))
}

fn checked_mul(left: usize, right: usize, context: &'static str) -> RewriteResult<usize> {
    left.checked_mul(right).ok_or_else(|| arithmetic(context))
}

fn require_exact_capacity<T>(
    value: &Vec<T>,
    expected: usize,
    resource: &'static str,
) -> RewriteResult<()> {
    if size_of::<T>() != 0 && value.capacity() != expected {
        return Err(RewriteError::Allocation {
            resource,
            amount: expected,
        });
    }
    Ok(())
}

fn preflight_root_text(source: &[u8], limits: RewriteLimits) -> RewriteResult<TextPreflight> {
    if source.len() > limits.max_message_bytes() {
        return Err(RewriteError::LimitExceeded {
            resource: "message bytes",
            observed: source.len(),
            limit: limits.max_message_bytes(),
        });
    }
    let mut occurrences = [0usize; MAX_KNOWN_ROOT_FIELD + 1];
    let mut fragments = 0usize;
    let mut text_bytes = 0usize;
    let mut utf16_len = 0usize;
    let mut field_count = 0usize;
    let mut fields = RawFields::new(source);
    while let Some(field) = fields.next()? {
        field_count = checked_add(field_count, 1, "root field count")?;
        enforce_limit("fields", field_count, limits.max_fields())?;
        let number =
            usize::try_from(field.number).map_err(|_error| arithmetic("root field number"))?;
        if number <= MAX_KNOWN_ROOT_FIELD && is_known_root_field(field.number) {
            occurrences[number] =
                checked_add(occurrences[number], 1, "known root field occurrences")?;
            validate_root_field_shape(field)?;
            if field.number != ROOT_TEXT_FIELD && occurrences[number] > 1 {
                return Err(invalid_owned(format!(
                    "singular TSWP storage field {} occurs {} times",
                    field.number, occurrences[number]
                )));
            }
        }
        if field.number == ROOT_TEXT_FIELD {
            let payload = field.canonical_payload(2, "TSWP storage text")?;
            let fragment_index = fragments;
            fragments = checked_add(fragments, 1, "text fragment count")?;
            text_bytes = checked_add(text_bytes, payload.len(), "aggregate text bytes")?;
            let fragment = std::str::from_utf8(payload).map_err(|error| {
                invalid_owned(format!(
                    "text fragment {fragment_index} is invalid UTF-8 at byte {}",
                    error.valid_up_to()
                ))
            })?;
            for character in fragment.chars() {
                utf16_len =
                    checked_add(utf16_len, character.len_utf16(), "preflight UTF-16 length")?;
            }
        }
    }
    enforce_limit("text fragments", fragments, limits.max_fragments())?;
    enforce_limit("text bytes", text_bytes, limits.max_text_bytes())?;
    let _utf16_u32 =
        u32::try_from(utf16_len).map_err(|_error| arithmetic("source UTF-16 u32 length"))?;
    Ok(TextPreflight {
        fragments,
        text_bytes,
        utf16_len,
    })
}

fn is_known_root_field(number: u32) -> bool {
    matches!(number, 1..=12 | 14..=28)
}

fn is_table_field(number: u32) -> bool {
    matches!(number, 5..=9 | 11..=12 | 14..=28)
}

fn validate_root_field_shape(field: RawField<'_>) -> RewriteResult<()> {
    match field.number {
        1 => {
            let _kind = field.canonical_varint("TSWP storage kind")?;
        },
        2 => {
            let _reference = field.canonical_payload(2, "TSWP storage style sheet")?;
        },
        3 => {
            let _text = field.canonical_payload(2, "TSWP storage text")?;
        },
        4 | 10 => {
            let value = field.canonical_varint("TSWP storage boolean")?;
            if value > 1 {
                return Err(invalid_owned(format!(
                    "TSWP storage boolean field {} has value {value}",
                    field.number
                )));
            }
        },
        number if is_table_field(number) => {
            let _table = field.canonical_payload(2, "TSWP storage table")?;
        },
        _ => {},
    }
    Ok(())
}

fn enforce_limit(resource: &'static str, observed: usize, limit: usize) -> RewriteResult<()> {
    if observed > limit {
        return Err(RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        });
    }
    Ok(())
}

fn decode_text_plan(
    source: &[u8],
    range: &Range<usize>,
    replacement: &str,
    preflight: TextPreflight,
    limits: RewriteLimits,
) -> RewriteResult<TextPlan> {
    let element_memory = checked_mul(
        preflight.fragments,
        size_of::<&str>(),
        "Buffa fragment metadata",
    )?;
    let options =
        text_storage_codec::DecodeOptions::new(limits.max_message_bytes(), 0, element_memory, 1);
    let view = text_storage_codec::decode_storage_text(source, options)
        .map_err(|error| RewriteError::Projection(error.to_string()))?;
    if view.len() != preflight.fragments {
        return Err(invalid_owned(format!(
            "Buffa returned {} fragments after raw preflight counted {}",
            view.len(),
            preflight.fragments
        )));
    }
    let (before_utf16_len, observed_text_bytes) =
        text_lengths(view.fragments(), limits.max_text_bytes())?;
    if observed_text_bytes != preflight.text_bytes {
        return Err(invalid("Buffa text length disagrees with raw preflight"));
    }
    if before_utf16_len != preflight.utf16_len {
        return Err(invalid("Buffa UTF-16 length disagrees with raw preflight"));
    }
    if range.start > range.end {
        return Err(RewriteError::ReversedRange {
            start: range.start,
            end: range.end,
        });
    }
    if range.end > before_utf16_len {
        return Err(RewriteError::RangeOutOfBounds {
            index: range.end,
            length: before_utf16_len,
        });
    }
    let start_byte = utf16_to_utf8(view.fragments(), range.start)?;
    let end_byte = utf16_to_utf8(view.fragments(), range.end)?;
    let replacement_units = replacement.encode_utf16().count();
    let removed_units = range
        .end
        .checked_sub(range.start)
        .ok_or_else(|| arithmetic("removed UTF-16 length"))?;
    let after_utf16_len = before_utf16_len
        .checked_sub(removed_units)
        .and_then(|length| length.checked_add(replacement_units))
        .ok_or_else(|| arithmetic("rewritten UTF-16 length"))?;
    let _after_u32 = u32::try_from(after_utf16_len)
        .map_err(|_error| arithmetic("rewritten UTF-16 u32 length"))?;
    let removed_bytes = end_byte
        .checked_sub(start_byte)
        .ok_or_else(|| arithmetic("removed UTF-8 bytes"))?;
    let result_bytes = preflight
        .text_bytes
        .checked_sub(removed_bytes)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| arithmetic("rewritten UTF-8 length"))?;
    enforce_limit("text bytes", result_bytes, limits.max_text_bytes())?;
    let no_op = selected_text_equals(
        view.fragments(),
        start_byte,
        end_byte,
        replacement.as_bytes(),
    );
    let anchor_fragment = if no_op {
        None
    } else {
        insertion_anchor(view.fragments(), start_byte, preflight.fragments)
    };
    Ok(TextPlan {
        before_utf16_len,
        after_utf16_len,
        start_byte,
        end_byte,
        replacement_units,
        result_bytes,
        anchor_fragment,
        no_op,
    })
}

fn text_lengths<'source>(
    fragments: impl ExactSizeIterator<Item = &'source str>,
    text_limit: usize,
) -> RewriteResult<(usize, usize)> {
    let mut utf16_len = 0usize;
    let mut utf8_len = 0usize;
    for fragment in fragments {
        utf8_len = checked_add(utf8_len, fragment.len(), "aggregate UTF-8 length")?;
        enforce_limit("text bytes", utf8_len, text_limit)?;
        for character in fragment.chars() {
            utf16_len = checked_add(utf16_len, character.len_utf16(), "aggregate UTF-16 length")?;
        }
    }
    let _utf16_u32 =
        u32::try_from(utf16_len).map_err(|_error| arithmetic("source UTF-16 u32 length"))?;
    Ok((utf16_len, utf8_len))
}

fn utf16_to_utf8<'source>(
    fragments: impl ExactSizeIterator<Item = &'source str>,
    target: usize,
) -> RewriteResult<usize> {
    let mut utf16_offset = 0usize;
    let mut utf8_offset = 0usize;
    if target == 0 {
        return Ok(0);
    }
    for fragment in fragments {
        for character in fragment.chars() {
            if utf16_offset == target {
                return Ok(utf8_offset);
            }
            utf16_offset =
                checked_add(utf16_offset, character.len_utf16(), "UTF-16 boundary scan")?;
            utf8_offset = checked_add(utf8_offset, character.len_utf8(), "UTF-8 boundary scan")?;
            if utf16_offset > target {
                return Err(RewriteError::SurrogateSplit { index: target });
            }
        }
    }
    if utf16_offset == target {
        Ok(utf8_offset)
    } else {
        Err(RewriteError::RangeOutOfBounds {
            index: target,
            length: utf16_offset,
        })
    }
}

fn selected_text_equals<'source>(
    fragments: impl ExactSizeIterator<Item = &'source str>,
    start_byte: usize,
    end_byte: usize,
    replacement: &[u8],
) -> bool {
    if end_byte.saturating_sub(start_byte) != replacement.len() {
        return false;
    }
    let mut global = 0usize;
    let mut compared = 0usize;
    for fragment in fragments {
        let fragment_end = global.saturating_add(fragment.len());
        let selected_start = start_byte.saturating_sub(global).min(fragment.len());
        let selected_end = end_byte.saturating_sub(global).min(fragment.len());
        if selected_start < selected_end {
            let count = selected_end - selected_start;
            let Some(expected) = replacement.get(compared..compared.saturating_add(count)) else {
                return false;
            };
            if &fragment.as_bytes()[selected_start..selected_end] != expected {
                return false;
            }
            compared = compared.saturating_add(count);
        }
        global = fragment_end;
        if global >= end_byte {
            break;
        }
    }
    compared == replacement.len()
}

fn insertion_anchor<'source>(
    fragments: impl ExactSizeIterator<Item = &'source str>,
    start_byte: usize,
    fragment_count: usize,
) -> Option<usize> {
    let mut global = 0usize;
    let mut last = None;
    let mut first = None;
    for (index, fragment) in fragments.enumerate() {
        first.get_or_insert(index);
        last = Some(index);
        let end = global.saturating_add(fragment.len());
        if start_byte < end {
            return Some(index);
        }
        global = end;
    }
    if start_byte == 0 && fragment_count != 0 {
        first
    } else {
        last
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum EntryKind {
    Object,
    Para,
    String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TablePolicy {
    RetainObject,
    RetainPara,
    RetainString,
    DropObject,
    DropEmptyObject,
    NormalizeObject,
    Overlapping,
}

impl TablePolicy {
    const fn for_field(field: u32) -> Option<Self> {
        match field {
            5 | 7 | 8 | 12 | 17 | 28 => Some(Self::RetainObject),
            6 | 14 | 24 => Some(Self::RetainPara),
            19 | 20 => Some(Self::RetainString),
            9 | 16 => Some(Self::DropEmptyObject),
            11 | 15 | 23 => Some(Self::NormalizeObject),
            18 | 21 | 22 | 27 => Some(Self::DropObject),
            25 | 26 => Some(Self::Overlapping),
            _ => None,
        }
    }

    const fn entry_kind(self) -> Option<EntryKind> {
        match self {
            Self::RetainObject
            | Self::DropObject
            | Self::DropEmptyObject
            | Self::NormalizeObject => Some(EntryKind::Object),
            Self::RetainPara => Some(EntryKind::Para),
            Self::RetainString => Some(EntryKind::String),
            Self::Overlapping => None,
        }
    }

    const fn retain_start(self) -> bool {
        matches!(
            self,
            Self::RetainObject | Self::RetainPara | Self::RetainString
        )
    }
}

#[derive(Debug, Default)]
struct Counters {
    fields: usize,
    table_entries: usize,
    references: usize,
    tree_bytes: usize,
    unknown_fields: usize,
}

impl Counters {
    fn enter_message(
        &mut self,
        bytes: usize,
        depth: usize,
        limits: RewriteLimits,
    ) -> RewriteResult<()> {
        enforce_limit("nesting", depth, limits.max_nesting())?;
        self.tree_bytes = checked_add(self.tree_bytes, bytes, "aggregate nested scan bytes")?;
        enforce_limit(
            "aggregate nested scan bytes",
            self.tree_bytes,
            limits.max_rewrite_work(),
        )
    }

    fn field(&mut self, limits: RewriteLimits) -> RewriteResult<()> {
        self.fields = checked_add(self.fields, 1, "aggregate field count")?;
        enforce_limit("fields", self.fields, limits.max_fields())
    }

    fn table_entry(&mut self, limits: RewriteLimits) -> RewriteResult<()> {
        self.table_entries = checked_add(self.table_entries, 1, "table entry count")?;
        enforce_limit(
            "table entries",
            self.table_entries,
            limits.max_table_entries(),
        )
    }

    fn reference(&mut self, limits: RewriteLimits) -> RewriteResult<()> {
        self.references = checked_add(self.references, 1, "object reference count")?;
        enforce_limit(
            "object references",
            self.references,
            limits.max_object_references(),
        )
    }

    fn unknown_field(&mut self) -> RewriteResult<()> {
        self.unknown_fields = checked_add(self.unknown_fields, 1, "unknown field count")?;
        Ok(())
    }
}

#[derive(Debug, Clone, Copy)]
struct BorrowedIndexEntry<'source> {
    index: u32,
    object: Option<u64>,
    index_field: RawField<'source>,
}

#[derive(Debug, Clone, Copy)]
struct OverlappingEntry<'source> {
    start: u32,
    length: u32,
    reference: u64,
    range_field: RawField<'source>,
    location_field: RawField<'source>,
}

fn validate_reference(
    data: &[u8],
    depth: usize,
    counters: &mut Counters,
    limits: RewriteLimits,
) -> RewriteResult<u64> {
    counters.enter_message(data.len(), depth, limits)?;
    let mut identifier = None;
    let mut deprecated_type = false;
    let mut external = false;
    let mut fields = RawFields::new(data);
    while let Some(field) = fields.next()? {
        counters.field(limits)?;
        match field.number {
            1 => {
                require_absent(identifier.is_some(), "TSP.Reference identifier")?;
                identifier = Some(field.canonical_varint("TSP.Reference identifier")?);
            },
            2 => {
                require_absent(deprecated_type, "TSP.Reference deprecated type")?;
                let _value = field.canonical_varint("TSP.Reference deprecated type")?;
                deprecated_type = true;
            },
            3 => {
                require_absent(external, "TSP.Reference external flag")?;
                let value = field.canonical_varint("TSP.Reference external flag")?;
                if value > 1 {
                    return Err(invalid("TSP.Reference external flag is not boolean"));
                }
                external = true;
            },
            _ => {},
        }
        if !matches!(field.number, 1..=3) {
            counters.unknown_field()?;
        }
    }
    let required_identifier =
        identifier.ok_or_else(|| invalid("TSP.Reference identifier is missing"))?;
    if required_identifier == 0 {
        return Err(invalid("TSP.Reference identifier is zero"));
    }
    counters.reference(limits)?;
    Ok(required_identifier)
}

fn inspect_reference(data: &[u8]) -> RewriteResult<u64> {
    let mut identifier = None;
    let mut fields = RawFields::new(data);
    while let Some(field) = fields.next()? {
        if field.number == 1 {
            require_absent(identifier.is_some(), "TSP.Reference identifier")?;
            identifier = Some(field.canonical_varint("TSP.Reference identifier")?);
        }
    }
    identifier.ok_or_else(|| invalid("TSP.Reference identifier is missing"))
}

fn validate_range<'source>(
    data: &'source [u8],
    depth: usize,
    counters: &mut Counters,
    limits: RewriteLimits,
) -> RewriteResult<(u32, u32, RawField<'source>)> {
    counters.enter_message(data.len(), depth, limits)?;
    let mut location = None;
    let mut length = None;
    let mut location_field = None;
    let mut fields = RawFields::new(data);
    while let Some(field) = fields.next()? {
        counters.field(limits)?;
        match field.number {
            1 => {
                require_absent(location.is_some(), "TSP.Range location")?;
                location = Some(to_u32(
                    field.canonical_varint("TSP.Range location")?,
                    "TSP.Range location",
                )?);
                location_field = Some(field);
            },
            2 => {
                require_absent(length.is_some(), "TSP.Range length")?;
                length = Some(to_u32(
                    field.canonical_varint("TSP.Range length")?,
                    "TSP.Range length",
                )?);
            },
            _ => {},
        }
        if !matches!(field.number, 1..=2) {
            counters.unknown_field()?;
        }
    }
    Ok((
        location.ok_or_else(|| invalid("TSP.Range location is missing"))?,
        length.ok_or_else(|| invalid("TSP.Range length is missing"))?,
        location_field.ok_or_else(|| invalid("TSP.Range location field is missing"))?,
    ))
}

fn inspect_range(data: &[u8]) -> RewriteResult<(u32, u32, RawField<'_>)> {
    let mut location = None;
    let mut length = None;
    let mut location_field = None;
    let mut fields = RawFields::new(data);
    while let Some(field) = fields.next()? {
        match field.number {
            1 => {
                require_absent(location.is_some(), "TSP.Range location")?;
                location = Some(to_u32(
                    field.canonical_varint("TSP.Range location")?,
                    "TSP.Range location",
                )?);
                location_field = Some(field);
            },
            2 => {
                require_absent(length.is_some(), "TSP.Range length")?;
                length = Some(to_u32(
                    field.canonical_varint("TSP.Range length")?,
                    "TSP.Range length",
                )?);
            },
            _ => {},
        }
    }
    Ok((
        location.ok_or_else(|| invalid("TSP.Range location is missing"))?,
        length.ok_or_else(|| invalid("TSP.Range length is missing"))?,
        location_field.ok_or_else(|| invalid("TSP.Range location field is missing"))?,
    ))
}

fn validate_index_entry<'source>(
    data: &'source [u8],
    kind: EntryKind,
    depth: usize,
    counters: &mut Counters,
    limits: RewriteLimits,
) -> RewriteResult<BorrowedIndexEntry<'source>> {
    counters.enter_message(data.len(), depth, limits)?;
    let mut index = None;
    let mut index_field = None;
    let mut object = None;
    let mut value_two = false;
    let mut value_three = false;
    let mut fields = RawFields::new(data);
    while let Some(field) = fields.next()? {
        counters.field(limits)?;
        match field.number {
            1 => {
                require_absent(index.is_some(), "table entry character index")?;
                index = Some(to_u32(
                    field.canonical_varint("table entry character index")?,
                    "table entry character index",
                )?);
                index_field = Some(field);
            },
            2 => match kind {
                EntryKind::Object => {
                    require_absent(value_two, "object table entry reference")?;
                    let payload = field.canonical_payload(2, "object table entry reference")?;
                    object = Some(validate_reference(payload, depth + 1, counters, limits)?);
                    value_two = true;
                },
                EntryKind::Para => {
                    require_absent(value_two, "paragraph table first value")?;
                    let _first = to_u32(
                        field.canonical_varint("paragraph table first value")?,
                        "paragraph table first value",
                    )?;
                    value_two = true;
                },
                EntryKind::String => {
                    require_absent(value_two, "string table value")?;
                    let payload = field.canonical_payload(2, "string table value")?;
                    std::str::from_utf8(payload)
                        .map_err(|_error| invalid("string table value is not UTF-8"))?;
                    value_two = true;
                },
            },
            3 if kind == EntryKind::Para => {
                require_absent(value_three, "paragraph table second value")?;
                let _second = to_u32(
                    field.canonical_varint("paragraph table second value")?,
                    "paragraph table second value",
                )?;
                value_three = true;
            },
            _ => {},
        }
        let known = match kind {
            EntryKind::Object | EntryKind::String => matches!(field.number, 1..=2),
            EntryKind::Para => matches!(field.number, 1..=3),
        };
        if !known {
            counters.unknown_field()?;
        }
    }
    if kind == EntryKind::Para && (!value_two || !value_three) {
        return Err(invalid("paragraph table entry is missing required values"));
    }
    Ok(BorrowedIndexEntry {
        index: index.ok_or_else(|| invalid("table entry character index is missing"))?,
        object,
        index_field: index_field
            .ok_or_else(|| invalid("table entry character index field is missing"))?,
    })
}

fn inspect_index_entry(data: &[u8], kind: EntryKind) -> RewriteResult<BorrowedIndexEntry<'_>> {
    let mut index = None;
    let mut index_field = None;
    let mut object = None;
    let mut fields = RawFields::new(data);
    while let Some(field) = fields.next()? {
        match field.number {
            1 => {
                index = Some(to_u32(
                    field.canonical_varint("table entry character index")?,
                    "table entry character index",
                )?);
                index_field = Some(field);
            },
            2 if kind == EntryKind::Object => {
                object = Some(inspect_reference(
                    field.canonical_payload(2, "object table entry reference")?,
                )?);
            },
            _ => {},
        }
    }
    Ok(BorrowedIndexEntry {
        index: index.ok_or_else(|| invalid("table entry character index is missing"))?,
        object,
        index_field: index_field
            .ok_or_else(|| invalid("table entry character index field is missing"))?,
    })
}

fn validate_overlapping_entry<'source>(
    data: &'source [u8],
    depth: usize,
    counters: &mut Counters,
    limits: RewriteLimits,
) -> RewriteResult<OverlappingEntry<'source>> {
    counters.enter_message(data.len(), depth, limits)?;
    let mut range = None;
    let mut reference = None;
    let mut fields = RawFields::new(data);
    while let Some(field) = fields.next()? {
        counters.field(limits)?;
        match field.number {
            1 => {
                require_absent(range.is_some(), "overlapping table range")?;
                let payload = field.canonical_payload(2, "overlapping table range")?;
                let (start, length, location_field) =
                    validate_range(payload, depth + 1, counters, limits)?;
                range = Some((start, length, field, location_field));
            },
            2 => {
                require_absent(reference.is_some(), "overlapping table reference")?;
                let payload = field.canonical_payload(2, "overlapping table reference")?;
                reference = Some(validate_reference(payload, depth + 1, counters, limits)?);
            },
            _ => {},
        }
        if !matches!(field.number, 1..=2) {
            counters.unknown_field()?;
        }
    }
    let (start, length, range_field, location_field) =
        range.ok_or_else(|| invalid("overlapping table range is missing"))?;
    Ok(OverlappingEntry {
        start,
        length,
        reference: reference.ok_or_else(|| invalid("overlapping table reference is missing"))?,
        range_field,
        location_field,
    })
}

fn inspect_overlapping_entry(data: &[u8]) -> RewriteResult<OverlappingEntry<'_>> {
    let mut range = None;
    let mut reference = None;
    let mut fields = RawFields::new(data);
    while let Some(field) = fields.next()? {
        match field.number {
            1 => {
                let payload = field.canonical_payload(2, "overlapping table range")?;
                let (start, length, location_field) = inspect_range(payload)?;
                range = Some((start, length, field, location_field));
            },
            2 => {
                reference = Some(inspect_reference(
                    field.canonical_payload(2, "overlapping table reference")?,
                )?);
            },
            _ => {},
        }
    }
    let (start, length, range_field, location_field) =
        range.ok_or_else(|| invalid("overlapping table range is missing"))?;
    Ok(OverlappingEntry {
        start,
        length,
        reference: reference.ok_or_else(|| invalid("overlapping table reference is missing"))?,
        range_field,
        location_field,
    })
}

fn require_absent(already_present: bool, label: &'static str) -> RewriteResult<()> {
    if already_present {
        Err(invalid_owned(format!("duplicate {label}")))
    } else {
        Ok(())
    }
}

fn to_u32(value: u64, label: &'static str) -> RewriteResult<u32> {
    u32::try_from(value).map_err(|_error| invalid_owned(format!("{label} exceeds u32")))
}

#[derive(Debug, Clone, Copy)]
struct TablePlan {
    keep: bool,
    changed: bool,
    output_len: usize,
    generated_entries: usize,
}

#[derive(Debug, Clone, Copy)]
struct RootPlan {
    output_len: usize,
    reference_occurrences: usize,
    tree_bytes: usize,
    generated_fields: usize,
    has_unknown_wire_fields: bool,
    expected_changed: bool,
    skip_table_mutation: bool,
}

fn validate_and_plan_root(
    source: &[u8],
    range: &Range<usize>,
    text: TextPlan,
    replacement: &str,
    behavior: RewriteBehavior,
    limits: RewriteLimits,
) -> RewriteResult<RootPlan> {
    let mut counters = Counters::default();
    counters.enter_message(source.len(), 1, limits)?;
    let mut output_len = 0usize;
    let mut generated_fields = 0usize;
    let mut fragment_index = 0usize;
    let mut fragment_offset = 0usize;
    let mut table_changed = false;
    let skip_table_mutation = match behavior {
        RewriteBehavior::PreserveOnEqualText => text.no_op,
        RewriteBehavior::ReplaceSelection => range.start == range.end && replacement.is_empty(),
    };
    let mut fields = RawFields::new(source);
    while let Some(field) = fields.next()? {
        counters.field(limits)?;
        if !is_known_root_field(field.number) {
            counters.unknown_field()?;
        }
        let contribution = match field.number {
            2 => {
                let payload = field.canonical_payload(2, "TSWP storage style sheet")?;
                let _identifier = validate_reference(payload, 2, &mut counters, limits)?;
                field.raw().len()
            },
            ROOT_TEXT_FIELD => {
                let payload = field.canonical_payload(2, "TSWP storage text")?;
                let action = text_action(
                    fragment_index,
                    fragment_offset,
                    payload.len(),
                    text,
                    replacement.len(),
                )?;
                fragment_index = checked_add(fragment_index, 1, "text fragment index")?;
                fragment_offset =
                    checked_add(fragment_offset, payload.len(), "text fragment byte offset")?;
                text_action_encoded_len(
                    action,
                    payload.len(),
                    replacement.len(),
                    field.raw().len(),
                )?
            },
            number if is_table_field(number) => {
                let payload = field.canonical_payload(2, "TSWP storage table")?;
                let policy = TablePolicy::for_field(number)
                    .ok_or_else(|| invalid("recognized table has no rewrite policy"))?;
                let plan = validate_and_plan_table(
                    payload,
                    number,
                    policy,
                    range,
                    text.replacement_units,
                    text.before_utf16_len,
                    skip_table_mutation,
                    &mut counters,
                    limits,
                )?;
                generated_fields = checked_add(
                    generated_fields,
                    plan.generated_entries,
                    "generated table fields",
                )?;
                table_changed |= plan.changed || !plan.keep;
                if plan.keep {
                    if plan.changed {
                        encoded_length_delimited_len(number, plan.output_len)?
                    } else {
                        field.raw().len()
                    }
                } else {
                    0
                }
            },
            _ => field.raw().len(),
        };
        output_len = checked_add(output_len, contribution, "rewritten storage bytes")?;
    }
    if fragment_index == 0 && !text.no_op && !replacement.is_empty() {
        output_len = checked_add(
            output_len,
            encoded_length_delimited_len(ROOT_TEXT_FIELD, replacement.len())?,
            "appended text field",
        )?;
        generated_fields = checked_add(generated_fields, 1, "generated text fields")?;
    }
    if fragment_offset != text.result_source_bytes(replacement.len())? {
        return Err(invalid("raw fragment bytes disagree with text plan"));
    }
    enforce_limit("output bytes", output_len, limits.max_output_bytes())?;
    let output_field_bound = checked_add(counters.fields, generated_fields, "output field bound")?;
    enforce_limit("fields", output_field_bound, limits.max_fields())?;
    Ok(RootPlan {
        output_len,
        reference_occurrences: counters.references,
        tree_bytes: counters.tree_bytes,
        generated_fields,
        has_unknown_wire_fields: counters.unknown_fields != 0,
        expected_changed: !text.no_op || table_changed,
        skip_table_mutation,
    })
}

impl TextPlan {
    fn result_source_bytes(self, replacement_bytes: usize) -> RewriteResult<usize> {
        let removed = self
            .end_byte
            .checked_sub(self.start_byte)
            .ok_or_else(|| arithmetic("text plan removed bytes"))?;
        self.result_bytes
            .checked_sub(replacement_bytes)
            .and_then(|length| length.checked_add(removed))
            .ok_or_else(|| arithmetic("text plan source bytes"))
    }
}

#[derive(Debug, Clone, Copy)]
enum TextAction {
    Raw,
    Drop,
    Rewrite {
        prefix: usize,
        suffix_start: usize,
        include_replacement: bool,
    },
}

impl TextAction {
    fn payload_len(self, fragment_len: usize, replacement_len: usize) -> RewriteResult<usize> {
        match self {
            Self::Raw => Ok(fragment_len),
            Self::Drop => Ok(0),
            Self::Rewrite {
                prefix,
                suffix_start,
                include_replacement,
            } => {
                let suffix = fragment_len
                    .checked_sub(suffix_start)
                    .ok_or_else(|| arithmetic("text fragment suffix"))?;
                let base = checked_add(prefix, suffix, "rewritten fragment payload")?;
                if include_replacement {
                    checked_add(base, replacement_len, "rewritten fragment replacement")
                } else {
                    Ok(base)
                }
            },
        }
    }
}

fn text_action(
    fragment_index: usize,
    fragment_start: usize,
    fragment_len: usize,
    text: TextPlan,
    replacement_len: usize,
) -> RewriteResult<TextAction> {
    if text.no_op {
        return Ok(TextAction::Raw);
    }
    let fragment_end = checked_add(fragment_start, fragment_len, "text fragment end")?;
    let action = if text.start_byte == text.end_byte {
        if text.anchor_fragment == Some(fragment_index) {
            TextAction::Rewrite {
                prefix: text.start_byte.saturating_sub(fragment_start),
                suffix_start: text.start_byte.saturating_sub(fragment_start),
                include_replacement: true,
            }
        } else {
            TextAction::Raw
        }
    } else if fragment_end <= text.start_byte || fragment_start >= text.end_byte {
        TextAction::Raw
    } else if text.anchor_fragment == Some(fragment_index) {
        let prefix = text.start_byte.saturating_sub(fragment_start);
        let suffix_start = text
            .end_byte
            .saturating_sub(fragment_start)
            .min(fragment_len);
        TextAction::Rewrite {
            prefix,
            suffix_start,
            include_replacement: true,
        }
    } else if fragment_end <= text.end_byte {
        TextAction::Drop
    } else {
        TextAction::Rewrite {
            prefix: 0,
            suffix_start: text.end_byte.saturating_sub(fragment_start),
            include_replacement: false,
        }
    };
    if let TextAction::Rewrite { .. } = action {
        let payload_len = action.payload_len(fragment_len, replacement_len)?;
        if payload_len == 0 {
            return Ok(TextAction::Drop);
        }
    }
    Ok(action)
}

fn text_action_encoded_len(
    action: TextAction,
    fragment_len: usize,
    replacement_len: usize,
    original_len: usize,
) -> RewriteResult<usize> {
    match action {
        TextAction::Raw => Ok(original_len),
        TextAction::Drop => Ok(0),
        TextAction::Rewrite { .. } => encoded_length_delimited_len(
            ROOT_TEXT_FIELD,
            action.payload_len(fragment_len, replacement_len)?,
        ),
    }
}

fn encoded_length_delimited_len(field: u32, payload_len: usize) -> RewriteResult<usize> {
    let key = u64::from(field)
        .checked_shl(3)
        .and_then(|value| value.checked_add(2))
        .ok_or_else(|| arithmetic("protobuf length-delimited key"))?;
    let payload_u64 =
        u64::try_from(payload_len).map_err(|_error| arithmetic("protobuf payload u64 length"))?;
    checked_add(
        checked_add(
            varint_len(key),
            varint_len(payload_u64),
            "encoded field framing",
        )?,
        payload_len,
        "encoded field payload",
    )
}

fn validate_and_plan_table(
    table: &[u8],
    storage_field: u32,
    policy: TablePolicy,
    range: &Range<usize>,
    replacement_units: usize,
    text_len: usize,
    no_op: bool,
    counters: &mut Counters,
    limits: RewriteLimits,
) -> RewriteResult<TablePlan> {
    validate_table_tree(table, policy, text_len, counters, limits)?;
    if no_op {
        return Ok(TablePlan {
            keep: true,
            changed: false,
            output_len: table.len(),
            generated_entries: 0,
        });
    }
    if policy == TablePolicy::NormalizeObject {
        validate_normalized_source(table)?;
    }
    if policy == TablePolicy::Overlapping {
        plan_overlapping_table(table, range, replacement_units)
    } else {
        plan_index_table(table, policy, range, replacement_units)
    }
    .map_err(|error| add_table_context(error, storage_field))
}

fn validate_table_tree(
    table: &[u8],
    policy: TablePolicy,
    text_len: usize,
    counters: &mut Counters,
    limits: RewriteLimits,
) -> RewriteResult<()> {
    counters.enter_message(table.len(), 2, limits)?;
    let mut previous_index = None;
    let mut fields = RawFields::new(table);
    while let Some(field) = fields.next()? {
        counters.field(limits)?;
        if field.number != TABLE_ENTRY_FIELD {
            counters.unknown_field()?;
            continue;
        }
        let entry = field.canonical_payload(2, "TSWP table entry")?;
        counters.table_entry(limits)?;
        if policy == TablePolicy::Overlapping {
            let inspected = validate_overlapping_entry(entry, 3, counters, limits)?;
            validate_native_range(inspected.start, inspected.length, text_len)?;
        } else {
            let kind = policy
                .entry_kind()
                .ok_or_else(|| invalid("index table has no entry kind"))?;
            let inspected = validate_index_entry(entry, kind, 3, counters, limits)?;
            validate_index_order(previous_index, inspected.index)?;
            validate_index_bound(inspected.index, text_len)?;
            previous_index = Some(inspected.index);
        }
    }
    Ok(())
}

fn add_table_context(source_error: RewriteError, field: u32) -> RewriteError {
    match source_error {
        RewriteError::InvalidFormat(message) => {
            invalid_owned(format!("storage table field {field}: {message}"))
        },
        unchanged @ (RewriteError::InvalidLimit { .. }
        | RewriteError::LimitExceeded { .. }
        | RewriteError::ReversedRange { .. }
        | RewriteError::RangeOutOfBounds { .. }
        | RewriteError::SurrogateSplit { .. }
        | RewriteError::ArithmeticOverflow { .. }
        | RewriteError::Projection(_)
        | RewriteError::Allocation { .. }) => unchanged,
    }
}

fn validate_index_order(previous: Option<u32>, current: u32) -> RewriteResult<()> {
    if previous.is_some_and(|value| current <= value) {
        return Err(invalid_owned(format!(
            "table character indexes are duplicate or unsorted at {current}"
        )));
    }
    Ok(())
}

fn validate_index_bound(index: u32, text_len: usize) -> RewriteResult<()> {
    let native_index = usize::try_from(index).map_err(|_error| arithmetic("table index usize"))?;
    if native_index > text_len {
        return Err(invalid_owned(format!(
            "table character index {native_index} exceeds UTF-16 text length {text_len}"
        )));
    }
    Ok(())
}

fn validate_native_range(start: u32, length: u32, text_len: usize) -> RewriteResult<()> {
    let native_start = usize::try_from(start).map_err(|_error| arithmetic("range start usize"))?;
    let native_length =
        usize::try_from(length).map_err(|_error| arithmetic("range length usize"))?;
    let end = checked_add(native_start, native_length, "native overlapping range")?;
    if end > text_len {
        return Err(invalid_owned(format!(
            "overlapping range {native_start}..{end} exceeds UTF-16 text length {text_len}"
        )));
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct Normalization {
    keep_from_ordinal: usize,
    force_zero_ordinal: Option<usize>,
    insert_sentinel: bool,
}

fn plan_index_table(
    table: &[u8],
    policy: TablePolicy,
    range: &Range<usize>,
    replacement_units: usize,
) -> RewriteResult<TablePlan> {
    let kind = policy
        .entry_kind()
        .ok_or_else(|| invalid("index table has no entry kind"))?;
    let normalization_option = if policy == TablePolicy::NormalizeObject {
        find_normalization(table, kind, range, replacement_units, policy.retain_start())?
    } else {
        Some(Normalization {
            keep_from_ordinal: 0,
            force_zero_ordinal: None,
            insert_sentinel: false,
        })
    };
    let Some(normalization) = normalization_option else {
        return Ok(TablePlan {
            keep: false,
            changed: true,
            output_len: 0,
            generated_entries: 0,
        });
    };
    let mut output_len = 0usize;
    let mut retained_entries = 0usize;
    let mut source_entries = 0usize;
    let mut unknown_table_fields = 0usize;
    let mut changed = normalization.keep_from_ordinal != 0 || normalization.insert_sentinel;
    let mut previous_adjusted = None;
    let mut ordinal = 0usize;
    let mut fields = RawFields::new(table);
    while let Some(field) = fields.next()? {
        if field.number != TABLE_ENTRY_FIELD {
            output_len = checked_add(output_len, field.raw().len(), "table unknown fields")?;
            unknown_table_fields =
                checked_add(unknown_table_fields, 1, "unknown table field count")?;
            continue;
        }
        source_entries = checked_add(source_entries, 1, "source table entries")?;
        let entry = field.canonical_payload(2, "TSWP table entry")?;
        let inspected = inspect_index_entry(entry, kind)?;
        let adjusted = adjust_index(
            inspected.index,
            range,
            replacement_units,
            policy.retain_start(),
        )?;
        let deduplicated = adjusted.is_some_and(|value| previous_adjusted == Some(value));
        if let Some(value) = adjusted
            && !deduplicated
        {
            previous_adjusted = Some(value);
        }
        let include =
            adjusted.is_some() && !deduplicated && ordinal >= normalization.keep_from_ordinal;
        if include {
            if normalization.insert_sentinel && ordinal == normalization.keep_from_ordinal {
                output_len = checked_add(output_len, 4, "normalized sentinel entry")?;
                retained_entries = checked_add(retained_entries, 1, "retained sentinel")?;
            }
            let adjusted_index =
                adjusted.ok_or_else(|| invalid("included index entry was dropped"))?;
            let final_index = if normalization.force_zero_ordinal == Some(ordinal) {
                0
            } else {
                adjusted_index
            };
            let encoded_entry_len = if final_index == inspected.index {
                field.raw().len()
            } else {
                let payload_len = rewritten_varint_message_len(
                    entry.len(),
                    inspected.index_field,
                    u64::from(final_index),
                )?;
                encoded_length_delimited_len(TABLE_ENTRY_FIELD, payload_len)?
            };
            output_len = checked_add(output_len, encoded_entry_len, "rewritten table entry")?;
            retained_entries = checked_add(retained_entries, 1, "retained table entries")?;
            changed |= final_index != inspected.index;
        } else {
            changed = true;
        }
        ordinal = checked_add(ordinal, 1, "table entry ordinal")?;
    }
    let removed_final_known_entry =
        source_entries != 0 && retained_entries == 0 && unknown_table_fields == 0;
    let keep = !(policy == TablePolicy::DropEmptyObject && removed_final_known_entry);
    if !keep {
        changed = true;
        output_len = 0;
    }
    Ok(TablePlan {
        keep,
        changed,
        output_len,
        generated_entries: usize::from(normalization.insert_sentinel && keep),
    })
}

fn find_normalization(
    table: &[u8],
    kind: EntryKind,
    range: &Range<usize>,
    replacement_units: usize,
    retain_start: bool,
) -> RewriteResult<Option<Normalization>> {
    let mut previous_adjusted = None;
    let mut previous_retained_ordinal = None;
    let mut ordinal = 0usize;
    let mut fields = RawFields::new(table);
    while let Some(field) = fields.next()? {
        if field.number != TABLE_ENTRY_FIELD {
            continue;
        }
        let entry = inspect_index_entry(field.canonical_payload(2, "TSWP table entry")?, kind)?;
        let adjusted = adjust_index(entry.index, range, replacement_units, retain_start)?;
        let retained = adjusted.is_some_and(|value| previous_adjusted != Some(value));
        if retained {
            let adjusted_index =
                adjusted.ok_or_else(|| invalid("retained normalization entry missing"))?;
            if entry.object.is_some() {
                return Ok(Some(match previous_retained_ordinal {
                    Some(previous) => Normalization {
                        keep_from_ordinal: previous,
                        force_zero_ordinal: Some(previous),
                        insert_sentinel: false,
                    },
                    None => Normalization {
                        keep_from_ordinal: ordinal,
                        force_zero_ordinal: None,
                        insert_sentinel: adjusted_index != 0,
                    },
                }));
            }
            previous_adjusted = Some(adjusted_index);
            previous_retained_ordinal = Some(ordinal);
        }
        ordinal = checked_add(ordinal, 1, "normalization entry ordinal")?;
    }
    Ok(None)
}

fn validate_normalized_source(table: &[u8]) -> RewriteResult<()> {
    let mut ordinal = 0usize;
    let mut first_index = None;
    let mut first_has_object = false;
    let mut first_object_ordinal = None;
    let mut fields = RawFields::new(table);
    while let Some(field) = fields.next()? {
        if field.number != TABLE_ENTRY_FIELD {
            continue;
        }
        let entry = inspect_index_entry(
            field.canonical_payload(2, "ranged object table entry")?,
            EntryKind::Object,
        )?;
        if ordinal == 0 {
            first_index = Some(entry.index);
            first_has_object = entry.object.is_some();
        }
        if first_object_ordinal.is_none() && entry.object.is_some() {
            first_object_ordinal = Some(ordinal);
        }
        ordinal = checked_add(ordinal, 1, "ranged object entry ordinal")?;
    }
    let Some(required_first_object_ordinal) = first_object_ordinal else {
        return Err(invalid(
            "ranged object table has no object-bearing entry and cannot be safely rewritten",
        ));
    };
    let normalized = match required_first_object_ordinal {
        0 => first_has_object && first_index == Some(0),
        1 => !first_has_object && first_index == Some(0),
        _ => false,
    };
    if !normalized {
        return Err(invalid(
            "ranged object table is not normalized; refusing global normalization",
        ));
    }
    Ok(())
}

fn plan_overlapping_table(
    table: &[u8],
    range: &Range<usize>,
    replacement_units: usize,
) -> RewriteResult<TablePlan> {
    let mut output_len = 0usize;
    let mut retained_entries = 0usize;
    let mut changed = false;
    let mut fields = RawFields::new(table);
    while let Some(field) = fields.next()? {
        if field.number != TABLE_ENTRY_FIELD {
            output_len = checked_add(output_len, field.raw().len(), "overlapping table unknown")?;
            continue;
        }
        let entry = field.canonical_payload(2, "overlapping table entry")?;
        let inspected = inspect_overlapping_entry(entry)?;
        match adjust_native_range(inspected.start, inspected.length, range, replacement_units)? {
            None => changed = true,
            Some(start) => {
                let encoded_entry_len = if start == inspected.start {
                    field.raw().len()
                } else {
                    let payload_len = rewritten_nested_location_len(entry, inspected, start)?;
                    encoded_length_delimited_len(TABLE_ENTRY_FIELD, payload_len)?
                };
                output_len = checked_add(output_len, encoded_entry_len, "overlapping table entry")?;
                retained_entries = checked_add(retained_entries, 1, "overlapping entries")?;
                changed |= start != inspected.start;
            },
        }
    }
    Ok(TablePlan {
        keep: true,
        changed,
        output_len,
        generated_entries: 0,
    })
}

fn adjust_index(
    index: u32,
    range: &Range<usize>,
    replacement_units: usize,
    retain_start: bool,
) -> RewriteResult<Option<u32>> {
    let index_usize = usize::try_from(index).map_err(|_error| arithmetic("table index usize"))?;
    if index_usize < range.start || (retain_start && index_usize == range.start) {
        return Ok(Some(index));
    }
    if index_usize < range.end {
        return Ok(None);
    }
    shift_index(index_usize, range, replacement_units).map(Some)
}

fn shift_index(index: usize, range: &Range<usize>, replacement_units: usize) -> RewriteResult<u32> {
    let removed = range
        .end
        .checked_sub(range.start)
        .ok_or_else(|| arithmetic("removed UTF-16 units"))?;
    let shifted = if replacement_units >= removed {
        index.checked_add(replacement_units - removed)
    } else {
        index.checked_sub(removed - replacement_units)
    }
    .ok_or_else(|| arithmetic("shifted table index"))?;
    u32::try_from(shifted).map_err(|_error| arithmetic("shifted table index u32"))
}

fn adjust_native_range(
    start: u32,
    length: u32,
    replacement: &Range<usize>,
    replacement_units: usize,
) -> RewriteResult<Option<u32>> {
    let start_usize = usize::try_from(start).map_err(|_error| arithmetic("range start usize"))?;
    let length_usize =
        usize::try_from(length).map_err(|_error| arithmetic("range length usize"))?;
    let end = checked_add(start_usize, length_usize, "overlapping range end")?;
    if end <= replacement.start {
        Ok(Some(start))
    } else if start_usize >= replacement.end {
        shift_index(start_usize, replacement, replacement_units).map(Some)
    } else {
        Ok(None)
    }
}

fn rewritten_varint_message_len(
    original_len: usize,
    original_field: RawField<'_>,
    replacement: u64,
) -> RewriteResult<usize> {
    original_len
        .checked_sub(original_field.payload().len())
        .and_then(|length| length.checked_add(varint_len(replacement)))
        .ok_or_else(|| arithmetic("rewritten varint message length"))
}

fn rewritten_nested_location_len(
    entry: &[u8],
    inspected: OverlappingEntry<'_>,
    replacement: u32,
) -> RewriteResult<usize> {
    let range_len = rewritten_varint_message_len(
        inspected.range_field.payload().len(),
        inspected.location_field,
        u64::from(replacement),
    )?;
    let without_range = entry
        .len()
        .checked_sub(inspected.range_field.raw().len())
        .ok_or_else(|| arithmetic("overlapping entry without range"))?;
    checked_add(
        without_range,
        encoded_length_delimited_len(inspected.range_field.number, range_len)?,
        "rewritten overlapping entry length",
    )
}

/// Rewrite one raw `TSWP.StorageArchive` without materializing its full Prost
/// model or surrendering unknown bytes to a generated encoder.
///
/// Text positions use UTF-16 code units, matching native iWork tables. An
/// equal selected/replacement string validates the complete known schema but
/// returns an exact source-byte copy without changing or normalizing tables.
/// During a real edit,
/// field-3 fragments wholly outside the edit remain byte-identical; insertion
/// at an interior fragment boundary has right affinity, while insertion at
/// end-of-text uses the final fragment.
///
/// # Errors
///
/// Returns a typed error for malformed or ambiguous known wire data, invalid
/// UTF-8/UTF-16 boundaries, checked arithmetic failure, Buffa disagreement,
/// an exceeded retained bound, or allocation failure. Unknown fields remain
/// opaque and byte-authoritative.
pub fn rewrite_storage_text_with_limits(
    source: &[u8],
    range: Range<usize>,
    replacement: &str,
    limits: RewriteLimits,
) -> RewriteResult<StorageRewrite> {
    rewrite_storage_text_with_behavior_and_limits(
        source,
        range,
        replacement,
        RewriteBehavior::PreserveOnEqualText,
        limits,
    )
}

/// Rewrite one raw storage under an explicit equal-text table behavior.
///
/// [`RewriteBehavior::ReplaceSelection`] mirrors the legacy editor's native
/// selection replacement: an equal nonempty text replacement may still remove
/// intersecting object/range attributes. Generic and Pages callers should use
/// [`RewriteBehavior::PreserveOnEqualText`] unless they intentionally need
/// that lifecycle behavior.
///
/// # Errors
///
/// Returns the same bounded, typed failures as
/// [`rewrite_storage_text_with_limits`].
pub fn rewrite_storage_text_with_behavior_and_limits(
    source: &[u8],
    range: Range<usize>,
    replacement: &str,
    behavior: RewriteBehavior,
    limits: RewriteLimits,
) -> RewriteResult<StorageRewrite> {
    let prepared = prepare_storage_text_rewrite_with_behavior_and_limits(
        source,
        range,
        replacement,
        behavior,
        limits,
    )?;
    let requirements = prepared.execution_requirements();
    prepared.execute(StorageRewriteExecutionLimits {
        max_output_bytes: requirements.output_bytes,
        max_retained_elements: requirements.retained_elements,
        max_retained_bytes: requirements.retained_bytes,
        max_peak_scratch_bytes: requirements.peak_scratch_bytes,
        max_allocations: requirements.allocations,
        max_work: requirements.work,
    })
}

/// Strictly plan one storage-text rewrite without allocating candidate output.
///
/// # Errors
///
/// Returns the same source, semantic, and planning-limit failures as
/// [`rewrite_storage_text_with_behavior_and_limits`].
pub fn prepare_storage_text_rewrite_with_behavior_and_limits<'source, 'replacement>(
    source: &'source [u8],
    range: Range<usize>,
    replacement: &'replacement str,
    behavior: RewriteBehavior,
    limits: RewriteLimits,
) -> RewriteResult<PreparedStorageRewrite<'source, 'replacement>> {
    let text_preflight = preflight_root_text(source, limits)?;
    let text = decode_text_plan(source, &range, replacement, text_preflight, limits)?;
    let root = validate_and_plan_root(source, &range, text, replacement, behavior, limits)?;
    let work = preflight_rewrite_work(
        source.len(),
        text_preflight,
        text,
        replacement,
        &root,
        limits,
    )?;
    let requirements = storage_execution_requirements(&root, work)?;
    Ok(PreparedStorageRewrite {
        source,
        range,
        replacement,
        text,
        root,
        limits,
        requirements,
    })
}

fn execute_prepared_storage_rewrite(
    prepared: PreparedStorageRewrite<'_, '_>,
) -> RewriteResult<StorageRewrite> {
    let PreparedStorageRewrite {
        source,
        range,
        replacement,
        text,
        root,
        limits,
        requirements,
    } = prepared;
    let references_before =
        collect_reference_occurrences(source, root.reference_occurrences, limits)?;
    let bytes = rewrite_root(
        source,
        &range,
        text,
        replacement,
        root.output_len,
        root.skip_table_mutation,
        limits,
    )?;
    let references_after =
        collect_reference_occurrences(&bytes, root.reference_occurrences, limits)?;
    let unique_references_before = unique_reference_pairs(&references_before)?;
    let unique_references_after = unique_reference_pairs(&references_after)?;
    let removed_by_field = sorted_difference(&unique_references_before, &unique_references_after)?;
    let object_reference_occurrences_before = occurrence_identifiers(&references_before)?;
    let object_reference_occurrences_after = occurrence_identifiers(&references_after)?;
    let object_references_before = aggregate_identifiers(&references_before)?;
    let object_references_after = aggregate_identifiers(&references_after)?;
    let removed_object_references =
        sorted_identifier_difference(&object_references_before, &object_references_after)?;
    let changed = bytes.as_slice() != source;
    if changed != root.expected_changed {
        return Err(invalid(
            "storage rewrite byte delta disagrees with semantic no-op plan",
        ));
    }
    let execution_report = storage_execution_report(
        &bytes,
        &object_references_before,
        &object_references_after,
        &object_reference_occurrences_before,
        &object_reference_occurrences_after,
        &removed_object_references,
        &removed_by_field,
        &references_before,
        &references_after,
        &unique_references_before,
        &unique_references_after,
        requirements.work,
    )?;
    Ok(StorageRewrite {
        bytes,
        before_utf16_len: text.before_utf16_len,
        after_utf16_len: text.after_utf16_len,
        object_references_before,
        object_references_after,
        object_reference_occurrences_before,
        object_reference_occurrences_after,
        removed_object_references,
        removed_object_references_by_field: removed_by_field,
        has_unknown_wire_fields: root.has_unknown_wire_fields,
        changed,
        execution_report,
    })
}

/// Validate the complete known `TSWP.StorageArchive` tree without producing a
/// rewritten byte buffer or materializing reference vectors.
///
/// The raw pass validates root cardinality, every positional table and entry,
/// required reference/range children, ordering, bounds, wire kinds, and all
/// retained limits. Buffa independently validates the borrowed field-3 text
/// projection after allocation-free raw bounds are established.
///
/// # Errors
///
/// Returns a typed bounded validation error. Unknown fields remain opaque and
/// are reported in [`StorageValidation`] rather than decoded.
pub fn validate_storage_with_limits(
    source: &[u8],
    limits: RewriteLimits,
) -> RewriteResult<StorageValidation> {
    let text_preflight = preflight_root_text(source, limits)?;
    let counters = validate_full_storage_tree(source, text_preflight.utf16_len, limits)?;
    let validation_work =
        preflight_validation_work(source.len(), text_preflight, &counters, limits)?;
    let empty_range = Range { start: 0, end: 0 };
    let text = decode_text_plan(source, &empty_range, "", text_preflight, limits)?;
    Ok(StorageValidation {
        utf8_len: text_preflight.text_bytes,
        utf16_len: text.before_utf16_len,
        fragments: text_preflight.fragments,
        fields: counters.fields,
        table_entries: counters.table_entries,
        reference_occurrences: counters.references,
        validation_work,
        has_unknown_wire_fields: counters.unknown_fields != 0,
    })
}

fn validate_full_storage_tree(
    source: &[u8],
    text_len: usize,
    limits: RewriteLimits,
) -> RewriteResult<Counters> {
    let mut counters = Counters::default();
    counters.enter_message(source.len(), 1, limits)?;
    let mut fields = RawFields::new(source);
    while let Some(field) = fields.next()? {
        counters.field(limits)?;
        if !is_known_root_field(field.number) {
            counters.unknown_field()?;
        }
        match field.number {
            2 => {
                let payload = field.canonical_payload(2, "TSWP storage style sheet")?;
                let _identifier = validate_reference(payload, 2, &mut counters, limits)?;
            },
            number if is_table_field(number) => {
                let payload = field.canonical_payload(2, "TSWP storage table")?;
                let policy = TablePolicy::for_field(number)
                    .ok_or_else(|| invalid("recognized table has no validation policy"))?;
                validate_table_tree(payload, policy, text_len, &mut counters, limits)
                    .map_err(|error| add_table_context(error, number))?;
            },
            _ => {},
        }
    }
    Ok(counters)
}

fn preflight_validation_work(
    source_len: usize,
    text: TextPreflight,
    counters: &Counters,
    limits: RewriteLimits,
) -> RewriteResult<usize> {
    let root_and_projection = checked_mul(source_len, 2, "validation root and Buffa work")?;
    let text_work = checked_mul(text.text_bytes, 6, "validation text work")?;
    let structural_work = checked_add(root_and_projection, counters.tree_bytes, "validation work")?;
    let aggregate_work = checked_add(structural_work, text_work, "validation work")?;
    enforce_limit("rewrite work", aggregate_work, limits.max_rewrite_work())?;
    Ok(aggregate_work)
}

fn preflight_rewrite_work(
    source_len: usize,
    text_preflight: TextPreflight,
    text: TextPlan,
    replacement: &str,
    root: &RootPlan,
    limits: RewriteLimits,
) -> RewriteResult<usize> {
    let root_scans = checked_mul(source_len, 2, "root preflight and Buffa work")?;
    let source_tree_scans = checked_mul(root.tree_bytes, 3, "source tree rewrite work")?;
    let output_tree_scan = checked_mul(
        root.output_len,
        MAX_SCHEMA_NESTING,
        "output reference scan work",
    )?;
    let text_scans = checked_mul(text_preflight.text_bytes, 6, "UTF text scan work")?;
    let output_copy = root.output_len;
    let mut work = checked_add(root_scans, source_tree_scans, "aggregate rewrite work")?;
    work = checked_add(work, output_tree_scan, "aggregate rewrite work")?;
    work = checked_add(work, text_scans, "aggregate rewrite work")?;
    work = checked_add(work, output_copy, "aggregate rewrite work")?;
    work = checked_add(work, replacement.len(), "aggregate rewrite work")?;
    work = checked_add(work, text.result_bytes, "aggregate rewrite work")?;
    work = checked_add(work, root.generated_fields, "aggregate rewrite work")?;
    enforce_limit("rewrite work", work, limits.max_rewrite_work())?;
    Ok(work)
}

fn storage_execution_requirements(
    root: &RootPlan,
    work: usize,
) -> RewriteResult<StorageRewriteExecutionRequirements> {
    let references = root.reference_occurrences;
    let retained_elements = root
        .output_len
        .checked_add(checked_mul(references, 6, "storage retained elements")?)
        .ok_or_else(|| arithmetic("storage retained elements"))?;
    let identifier_bytes = checked_mul(
        checked_mul(references, 5, "storage retained identifiers")?,
        size_of::<u64>(),
        "storage retained identifier bytes",
    )?;
    let provenance_bytes = checked_mul(
        references,
        size_of::<RemovedObjectReference>(),
        "storage retained provenance bytes",
    )?;
    let retained_bytes = root
        .output_len
        .checked_add(identifier_bytes)
        .and_then(|bytes| bytes.checked_add(provenance_bytes))
        .ok_or_else(|| arithmetic("storage retained bytes"))?;
    let intermediate = checked_mul(
        checked_mul(references, 4, "storage reference intermediates")?,
        size_of::<RemovedObjectReference>(),
        "storage reference intermediate bytes",
    )?;
    let peak_scratch_bytes = retained_bytes
        .checked_add(intermediate)
        .ok_or_else(|| arithmetic("storage execution peak"))?;
    let allocations = usize::from(root.output_len != 0)
        .checked_add(if references == 0 { 0 } else { 10 })
        .ok_or_else(|| arithmetic("storage execution allocations"))?;
    Ok(StorageRewriteExecutionRequirements {
        output_bytes: root.output_len,
        retained_elements,
        retained_bytes,
        peak_scratch_bytes,
        allocations,
        work,
        reference_occurrences: references,
    })
}

fn preflight_storage_execution(
    requirements: StorageRewriteExecutionRequirements,
    limits: StorageRewriteExecutionLimits,
) -> RewriteResult<()> {
    for (resource, observed, limit) in [
        (
            "output bytes",
            requirements.output_bytes,
            limits.max_output_bytes,
        ),
        (
            "retained elements",
            requirements.retained_elements,
            limits.max_retained_elements,
        ),
        (
            "retained bytes",
            requirements.retained_bytes,
            limits.max_retained_bytes,
        ),
        (
            "peak scratch bytes",
            requirements.peak_scratch_bytes,
            limits.max_peak_scratch_bytes,
        ),
        (
            "allocations",
            requirements.allocations,
            limits.max_allocations,
        ),
        ("rewrite work", requirements.work, limits.max_work),
    ] {
        enforce_limit(resource, observed, limit)?;
    }
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "each retained and intermediate storage vector remains explicit"
)]
fn storage_execution_report(
    bytes: &Vec<u8>,
    object_references_before: &Vec<u64>,
    object_references_after: &Vec<u64>,
    object_reference_occurrences_before: &Vec<u64>,
    object_reference_occurrences_after: &Vec<u64>,
    removed_object_references: &Vec<u64>,
    removed_by_field: &Vec<RemovedObjectReference>,
    references_before: &Vec<RemovedObjectReference>,
    references_after: &Vec<RemovedObjectReference>,
    unique_references_before: &Vec<RemovedObjectReference>,
    unique_references_after: &Vec<RemovedObjectReference>,
    work: usize,
) -> RewriteResult<StorageRewriteExecutionReport> {
    let identifier_elements = object_references_before
        .capacity()
        .checked_add(object_references_after.capacity())
        .and_then(|value| value.checked_add(object_reference_occurrences_before.capacity()))
        .and_then(|value| value.checked_add(object_reference_occurrences_after.capacity()))
        .and_then(|value| value.checked_add(removed_object_references.capacity()))
        .ok_or_else(|| arithmetic("storage actual identifier elements"))?;
    let retained_elements = bytes
        .capacity()
        .checked_add(identifier_elements)
        .and_then(|value| value.checked_add(removed_by_field.capacity()))
        .ok_or_else(|| arithmetic("storage actual retained elements"))?;
    let retained_bytes = bytes
        .capacity()
        .checked_add(checked_mul(
            identifier_elements,
            size_of::<u64>(),
            "storage actual identifier bytes",
        )?)
        .and_then(|value| {
            removed_by_field
                .capacity()
                .checked_mul(size_of::<RemovedObjectReference>())
                .and_then(|bytes| value.checked_add(bytes))
        })
        .ok_or_else(|| arithmetic("storage actual retained bytes"))?;
    let intermediate_elements = references_before
        .capacity()
        .checked_add(references_after.capacity())
        .and_then(|value| value.checked_add(unique_references_before.capacity()))
        .and_then(|value| value.checked_add(unique_references_after.capacity()))
        .ok_or_else(|| arithmetic("storage actual intermediate elements"))?;
    let peak_scratch_bytes = retained_bytes
        .checked_add(checked_mul(
            intermediate_elements,
            size_of::<RemovedObjectReference>(),
            "storage actual intermediate bytes",
        )?)
        .ok_or_else(|| arithmetic("storage actual peak scratch"))?;
    let allocations = usize::from(bytes.capacity() != 0)
        .checked_add(usize::from(references_before.capacity() != 0))
        .and_then(|value| value.checked_add(usize::from(references_after.capacity() != 0)))
        .and_then(|value| value.checked_add(usize::from(unique_references_before.capacity() != 0)))
        .and_then(|value| value.checked_add(usize::from(unique_references_after.capacity() != 0)))
        .and_then(|value| value.checked_add(usize::from(removed_by_field.capacity() != 0)))
        .and_then(|value| {
            value.checked_add(usize::from(
                object_reference_occurrences_before.capacity() != 0,
            ))
        })
        .and_then(|value| {
            value.checked_add(usize::from(
                object_reference_occurrences_after.capacity() != 0,
            ))
        })
        .and_then(|value| value.checked_add(usize::from(object_references_before.capacity() != 0)))
        .and_then(|value| value.checked_add(usize::from(object_references_after.capacity() != 0)))
        .and_then(|value| value.checked_add(usize::from(removed_object_references.capacity() != 0)))
        .ok_or_else(|| arithmetic("storage actual allocations"))?;
    Ok(StorageRewriteExecutionReport {
        retained_elements,
        retained_bytes,
        peak_scratch_bytes,
        allocations,
        work,
    })
}

fn rewrite_root(
    source: &[u8],
    range: &Range<usize>,
    text: TextPlan,
    replacement: &str,
    output_len: usize,
    skip_table_mutation: bool,
    limits: RewriteLimits,
) -> RewriteResult<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_allocation| RewriteError::Allocation {
            resource: "rewritten TSWP storage",
            amount: output_len,
        })?;
    require_exact_capacity(&output, output_len, "rewritten TSWP storage")?;
    let mut fragment_index = 0usize;
    let mut fragment_offset = 0usize;
    let mut fields = RawFields::new(source);
    while let Some(field) = fields.next()? {
        match field.number {
            ROOT_TEXT_FIELD => {
                let payload = field.canonical_payload(2, "TSWP storage text")?;
                let action = text_action(
                    fragment_index,
                    fragment_offset,
                    payload.len(),
                    text,
                    replacement.len(),
                )?;
                write_text_action(&mut output, field, action, replacement)?;
                fragment_index = checked_add(fragment_index, 1, "text fragment index")?;
                fragment_offset =
                    checked_add(fragment_offset, payload.len(), "text fragment offset")?;
            },
            number if is_table_field(number) => {
                let table = field.canonical_payload(2, "TSWP storage table")?;
                let policy = TablePolicy::for_field(number)
                    .ok_or_else(|| invalid("recognized table has no rewrite policy"))?;
                if skip_table_mutation {
                    output.extend_from_slice(field.raw());
                    continue;
                }
                let plan = if policy == TablePolicy::Overlapping {
                    plan_overlapping_table(table, range, text.replacement_units)?
                } else {
                    plan_index_table(table, policy, range, text.replacement_units)?
                };
                if !plan.keep {
                    continue;
                }
                if !plan.changed {
                    output.extend_from_slice(field.raw());
                    continue;
                }
                write_length_delimited_header(&mut output, number, plan.output_len)?;
                if policy == TablePolicy::Overlapping {
                    write_overlapping_table(&mut output, table, range, text.replacement_units)?;
                } else {
                    write_index_table(&mut output, table, policy, range, text.replacement_units)?;
                }
            },
            _ => output.extend_from_slice(field.raw()),
        }
    }
    if fragment_index == 0 && !text.no_op && !replacement.is_empty() {
        write_length_delimited_header(&mut output, ROOT_TEXT_FIELD, replacement.len())?;
        output.extend_from_slice(replacement.as_bytes());
    }
    if output.len() != output_len {
        return Err(invalid_owned(format!(
            "rewritten storage length {} disagrees with preflight {output_len}",
            output.len()
        )));
    }
    enforce_limit("output bytes", output.len(), limits.max_output_bytes())?;
    Ok(output)
}

fn write_text_action(
    output: &mut Vec<u8>,
    field: RawField<'_>,
    action: TextAction,
    replacement: &str,
) -> RewriteResult<()> {
    match action {
        TextAction::Raw => output.extend_from_slice(field.raw()),
        TextAction::Drop => {},
        TextAction::Rewrite {
            prefix,
            suffix_start,
            include_replacement,
        } => {
            let payload = field.payload();
            let payload_len = action.payload_len(payload.len(), replacement.len())?;
            write_length_delimited_header(output, ROOT_TEXT_FIELD, payload_len)?;
            output.extend_from_slice(
                payload
                    .get(..prefix)
                    .ok_or_else(|| invalid("text prefix exceeds its fragment"))?,
            );
            if include_replacement {
                output.extend_from_slice(replacement.as_bytes());
            }
            output.extend_from_slice(
                payload
                    .get(suffix_start..)
                    .ok_or_else(|| invalid("text suffix exceeds its fragment"))?,
            );
        },
    }
    Ok(())
}

fn write_length_delimited_header(
    output: &mut Vec<u8>,
    field: u32,
    payload_len: usize,
) -> RewriteResult<()> {
    let key = u64::from(field)
        .checked_shl(3)
        .and_then(|value| value.checked_add(2))
        .ok_or_else(|| arithmetic("protobuf length-delimited key"))?;
    let payload_length_u64 =
        u64::try_from(payload_len).map_err(|_error| arithmetic("protobuf payload u64 length"))?;
    encode_varint_into(output, key);
    encode_varint_into(output, payload_length_u64);
    Ok(())
}

fn write_index_table(
    output: &mut Vec<u8>,
    table: &[u8],
    policy: TablePolicy,
    range: &Range<usize>,
    replacement_units: usize,
) -> RewriteResult<()> {
    let kind = policy
        .entry_kind()
        .ok_or_else(|| invalid("index table has no entry kind"))?;
    let normalization = if policy == TablePolicy::NormalizeObject {
        find_normalization(table, kind, range, replacement_units, policy.retain_start())?
            .ok_or_else(|| invalid("removed normalized table reached writer"))?
    } else {
        Normalization {
            keep_from_ordinal: 0,
            force_zero_ordinal: None,
            insert_sentinel: false,
        }
    };
    let mut previous_adjusted = None;
    let mut ordinal = 0usize;
    let mut fields = RawFields::new(table);
    while let Some(field) = fields.next()? {
        if field.number != TABLE_ENTRY_FIELD {
            output.extend_from_slice(field.raw());
            continue;
        }
        let entry = field.canonical_payload(2, "TSWP table entry")?;
        let inspected = inspect_index_entry(entry, kind)?;
        let adjusted = adjust_index(
            inspected.index,
            range,
            replacement_units,
            policy.retain_start(),
        )?;
        let deduplicated = adjusted.is_some_and(|value| previous_adjusted == Some(value));
        if let Some(value) = adjusted
            && !deduplicated
        {
            previous_adjusted = Some(value);
        }
        if adjusted.is_some() && !deduplicated && ordinal >= normalization.keep_from_ordinal {
            if normalization.insert_sentinel && ordinal == normalization.keep_from_ordinal {
                output.extend_from_slice(&[0x0a, 0x02, 0x08, 0x00]);
            }
            let adjusted_index =
                adjusted.ok_or_else(|| invalid("included index entry was dropped"))?;
            let final_index = if normalization.force_zero_ordinal == Some(ordinal) {
                0
            } else {
                adjusted_index
            };
            if final_index == inspected.index {
                output.extend_from_slice(field.raw());
            } else {
                let payload_len = rewritten_varint_message_len(
                    entry.len(),
                    inspected.index_field,
                    u64::from(final_index),
                )?;
                write_length_delimited_header(output, TABLE_ENTRY_FIELD, payload_len)?;
                write_varint_message(output, entry, inspected.index_field, u64::from(final_index))?;
            }
        }
        ordinal = checked_add(ordinal, 1, "table entry ordinal")?;
    }
    Ok(())
}

fn write_overlapping_table(
    output: &mut Vec<u8>,
    table: &[u8],
    replacement: &Range<usize>,
    replacement_units: usize,
) -> RewriteResult<()> {
    let mut fields = RawFields::new(table);
    while let Some(field) = fields.next()? {
        if field.number != TABLE_ENTRY_FIELD {
            output.extend_from_slice(field.raw());
            continue;
        }
        let entry = field.canonical_payload(2, "overlapping table entry")?;
        let inspected = inspect_overlapping_entry(entry)?;
        let Some(start) = adjust_native_range(
            inspected.start,
            inspected.length,
            replacement,
            replacement_units,
        )?
        else {
            continue;
        };
        if start == inspected.start {
            output.extend_from_slice(field.raw());
            continue;
        }
        let payload_len = rewritten_nested_location_len(entry, inspected, start)?;
        write_length_delimited_header(output, TABLE_ENTRY_FIELD, payload_len)?;
        write_nested_location(output, entry, inspected, start)?;
    }
    Ok(())
}

fn write_varint_message(
    output: &mut Vec<u8>,
    message: &[u8],
    target: RawField<'_>,
    replacement: u64,
) -> RewriteResult<()> {
    let prefix = message
        .get(..target.start)
        .ok_or_else(|| invalid("varint target prefix exceeds message"))?;
    let suffix = message
        .get(target.end..)
        .ok_or_else(|| invalid("varint target suffix exceeds message"))?;
    output.extend_from_slice(prefix);
    output.extend_from_slice(target.key());
    encode_varint_into(output, replacement);
    output.extend_from_slice(suffix);
    Ok(())
}

fn write_nested_location(
    output: &mut Vec<u8>,
    entry: &[u8],
    inspected: OverlappingEntry<'_>,
    replacement: u32,
) -> RewriteResult<()> {
    output.extend_from_slice(
        entry
            .get(..inspected.range_field.start)
            .ok_or_else(|| invalid("range prefix exceeds overlapping entry"))?,
    );
    let range = inspected.range_field.payload();
    let range_len = rewritten_varint_message_len(
        range.len(),
        inspected.location_field,
        u64::from(replacement),
    )?;
    write_length_delimited_header(output, inspected.range_field.number, range_len)?;
    write_varint_message(
        output,
        range,
        inspected.location_field,
        u64::from(replacement),
    )?;
    output.extend_from_slice(
        entry
            .get(inspected.range_field.end..)
            .ok_or_else(|| invalid("range suffix exceeds overlapping entry"))?,
    );
    Ok(())
}

fn collect_reference_occurrences(
    source: &[u8],
    capacity: usize,
    limits: RewriteLimits,
) -> RewriteResult<Vec<RemovedObjectReference>> {
    let mut references = Vec::new();
    references
        .try_reserve_exact(capacity)
        .map_err(|_allocation| RewriteError::Allocation {
            resource: "TSWP object-reference occurrences",
            amount: capacity,
        })?;
    require_exact_capacity(&references, capacity, "TSWP reference occurrences")?;
    let mut fields = RawFields::new(source);
    while let Some(field) = fields.next()? {
        match field.number {
            2 => push_reference(
                &mut references,
                2,
                inspect_reference(field.canonical_payload(2, "TSWP storage style sheet")?)?,
                limits,
            )?,
            number if is_table_field(number) => {
                let policy = TablePolicy::for_field(number)
                    .ok_or_else(|| invalid("recognized table has no reference policy"))?;
                collect_table_references(
                    field.canonical_payload(2, "TSWP storage table")?,
                    number,
                    policy,
                    &mut references,
                    limits,
                )?;
            },
            _ => {},
        }
    }
    references.sort_unstable();
    Ok(references)
}

fn collect_table_references(
    table: &[u8],
    storage_field: u32,
    policy: TablePolicy,
    references: &mut Vec<RemovedObjectReference>,
    limits: RewriteLimits,
) -> RewriteResult<()> {
    let Some(kind) = policy.entry_kind() else {
        let mut fields = RawFields::new(table);
        while let Some(field) = fields.next()? {
            if field.number == TABLE_ENTRY_FIELD {
                let entry = inspect_overlapping_entry(
                    field.canonical_payload(2, "overlapping table entry")?,
                )?;
                push_reference(references, storage_field, entry.reference, limits)?;
            }
        }
        return Ok(());
    };
    if kind != EntryKind::Object {
        return Ok(());
    }
    let mut fields = RawFields::new(table);
    while let Some(field) = fields.next()? {
        if field.number == TABLE_ENTRY_FIELD {
            let entry = inspect_index_entry(
                field.canonical_payload(2, "object table entry")?,
                EntryKind::Object,
            )?;
            if let Some(identifier) = entry.object {
                push_reference(references, storage_field, identifier, limits)?;
            }
        }
    }
    Ok(())
}

fn push_reference(
    references: &mut Vec<RemovedObjectReference>,
    storage_field_number: u32,
    identifier: u64,
    limits: RewriteLimits,
) -> RewriteResult<()> {
    if references.len() >= limits.max_object_references() {
        return Err(RewriteError::LimitExceeded {
            resource: "object references",
            observed: references.len().saturating_add(1),
            limit: limits.max_object_references(),
        });
    }
    references.push(RemovedObjectReference {
        storage_field_number,
        identifier,
    });
    Ok(())
}

fn aggregate_identifiers(references: &[RemovedObjectReference]) -> RewriteResult<Vec<u64>> {
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(references.len())
        .map_err(|_allocation| RewriteError::Allocation {
            resource: "deduplicated TSWP object references",
            amount: references.len(),
        })?;
    require_exact_capacity(
        &identifiers,
        references.len(),
        "deduplicated TSWP object references",
    )?;
    identifiers.extend(references.iter().map(|reference| reference.identifier));
    identifiers.sort_unstable();
    identifiers.dedup();
    Ok(identifiers)
}

fn occurrence_identifiers(references: &[RemovedObjectReference]) -> RewriteResult<Vec<u64>> {
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(references.len())
        .map_err(|_allocation| RewriteError::Allocation {
            resource: "TSWP object-reference occurrence identifiers",
            amount: references.len(),
        })?;
    require_exact_capacity(
        &identifiers,
        references.len(),
        "TSWP object-reference occurrence identifiers",
    )?;
    identifiers.extend(references.iter().map(|reference| reference.identifier));
    identifiers.sort_unstable();
    Ok(identifiers)
}

fn unique_reference_pairs(
    references: &[RemovedObjectReference],
) -> RewriteResult<Vec<RemovedObjectReference>> {
    let mut unique = Vec::new();
    unique
        .try_reserve_exact(references.len())
        .map_err(|_allocation| RewriteError::Allocation {
            resource: "deduplicated TSWP reference provenance",
            amount: references.len(),
        })?;
    require_exact_capacity(
        &unique,
        references.len(),
        "deduplicated TSWP reference provenance",
    )?;
    unique.extend_from_slice(references);
    unique.dedup();
    Ok(unique)
}

fn sorted_difference(
    before: &[RemovedObjectReference],
    after: &[RemovedObjectReference],
) -> RewriteResult<Vec<RemovedObjectReference>> {
    let mut removed = Vec::new();
    removed
        .try_reserve_exact(before.len())
        .map_err(|_allocation| RewriteError::Allocation {
            resource: "removed TSWP reference provenance",
            amount: before.len(),
        })?;
    require_exact_capacity(&removed, before.len(), "removed TSWP reference provenance")?;
    let mut after_index = 0usize;
    for &candidate in before {
        while after
            .get(after_index)
            .is_some_and(|value| *value < candidate)
        {
            after_index = after_index.saturating_add(1);
        }
        if after.get(after_index) != Some(&candidate) {
            removed.push(candidate);
        }
    }
    Ok(removed)
}

fn sorted_identifier_difference(before: &[u64], after: &[u64]) -> RewriteResult<Vec<u64>> {
    let mut removed = Vec::new();
    removed
        .try_reserve_exact(before.len())
        .map_err(|_allocation| RewriteError::Allocation {
            resource: "removed TSWP object references",
            amount: before.len(),
        })?;
    require_exact_capacity(&removed, before.len(), "removed TSWP object references")?;
    let mut after_index = 0usize;
    for &candidate in before {
        while after
            .get(after_index)
            .is_some_and(|value| *value < candidate)
        {
            after_index = after_index.saturating_add(1);
        }
        if after.get(after_index) != Some(&candidate) {
            removed.push(candidate);
        }
    }
    Ok(removed)
}

#[cfg(test)]
#[allow(
    clippy::ref_option,
    clippy::shadow_reuse,
    reason = "Compact Prost-oracle helpers mirror generated optional field names"
)]
mod tests {
    use super::*;
    use litchi_iwa_protos::{
        tsp::{Range as NativeRange, Reference},
        tswp::{
            ObjectAttributeTable, OverlappingFieldAttributeTable, ParaDataAttributeTable,
            StorageArchive, StringAttributeTable, object_attribute_table::ObjectAttribute,
            overlapping_field_attribute_table::OverlappingFieldAttribute,
            para_data_attribute_table::ParaDataAttribute, string_attribute_table::StringAttribute,
        },
    };
    use prost::Message as _;

    fn reference(identifier: u64) -> Reference {
        Reference {
            identifier,
            deprecated_type: None,
            deprecated_is_external: None,
        }
    }

    fn object_table(entries: &[(u32, Option<u64>)]) -> ObjectAttributeTable {
        ObjectAttributeTable {
            entries: entries
                .iter()
                .map(|&(character_index, identifier)| ObjectAttribute {
                    character_index,
                    object: identifier.map(reference),
                })
                .collect(),
        }
    }

    fn para_table(indices: &[u32]) -> ParaDataAttributeTable {
        ParaDataAttributeTable {
            entries: indices
                .iter()
                .copied()
                .map(|character_index| ParaDataAttribute {
                    character_index,
                    first: 7,
                    second: 9,
                })
                .collect(),
        }
    }

    fn string_table(indices: &[u32]) -> StringAttributeTable {
        StringAttributeTable {
            entries: indices
                .iter()
                .copied()
                .map(|character_index| StringAttribute {
                    character_index,
                    object: Some("en".to_owned()),
                })
                .collect(),
        }
    }

    fn overlapping_table(base: u64) -> OverlappingFieldAttributeTable {
        OverlappingFieldAttributeTable {
            entries: vec![
                overlapping(0, 1, base),
                overlapping(1, 2, base + 1),
                overlapping(3, 1, base + 2),
            ],
        }
    }

    fn overlapping(start: u32, length: u32, identifier: u64) -> OverlappingFieldAttribute {
        OverlappingFieldAttribute {
            range: NativeRange {
                location: start,
                length,
            },
            field: reference(identifier),
        }
    }

    fn decode(bytes: &[u8]) -> StorageArchive {
        StorageArchive::decode(bytes)
            .unwrap_or_else(|error| panic!("test storage should decode through Prost: {error}"))
    }

    fn rewrite(source: &[u8], range: Range<usize>, replacement: &str) -> StorageRewrite {
        rewrite_storage_text_with_limits(source, range, replacement, RewriteLimits::default())
            .unwrap_or_else(|error| panic!("test storage should rewrite: {error}"))
    }

    fn raw_length_delimited(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::new();
        write_length_delimited_header(&mut bytes, number, payload.len())
            .unwrap_or_else(|error| panic!("test field should encode: {error}"));
        bytes.extend_from_slice(payload);
        bytes
    }

    fn raw_varint(number: u32, value: u64) -> Vec<u8> {
        let mut bytes = Vec::new();
        encode_varint_into(&mut bytes, u64::from(number) << 3);
        encode_varint_into(&mut bytes, value);
        bytes
    }

    fn root_field(source: &[u8], number: u32) -> Vec<u8> {
        let mut fields = RawFields::new(source);
        while let Some(field) = fields
            .next()
            .unwrap_or_else(|error| panic!("test root should parse: {error}"))
        {
            if field.number == number {
                return field.raw().to_vec();
            }
        }
        panic!("test root field {number} should exist")
    }

    fn object_indices(table: &Option<ObjectAttributeTable>) -> Vec<u32> {
        table
            .as_ref()
            .map(|table| {
                table
                    .entries
                    .iter()
                    .map(|entry| entry.character_index)
                    .collect()
            })
            .unwrap_or_default()
    }

    fn para_indices(table: &Option<ParaDataAttributeTable>) -> Vec<u32> {
        table
            .as_ref()
            .map(|table| {
                table
                    .entries
                    .iter()
                    .map(|entry| entry.character_index)
                    .collect()
            })
            .unwrap_or_default()
    }

    fn string_indices(table: &Option<StringAttributeTable>) -> Vec<u32> {
        table
            .as_ref()
            .map(|table| {
                table
                    .entries
                    .iter()
                    .map(|entry| entry.character_index)
                    .collect()
            })
            .unwrap_or_default()
    }

    #[test]
    fn exact_semantic_no_op_keeps_fragments_tables_and_unknowns() {
        let mut source = raw_length_delimited(3, b"A");
        source.extend_from_slice(&raw_varint(99, 7));
        source.extend_from_slice(&raw_length_delimited(3, "😀B".as_bytes()));
        source.extend_from_slice(&raw_length_delimited(
            5,
            &object_table(&[(0, Some(11)), (3, Some(12))]).encode_to_vec(),
        ));

        let result = rewrite(&source, 1..3, "😀");

        assert!(!result.changed());
        assert_eq!(result.bytes(), source);
        assert_eq!(result.before_utf16_len(), 4);
        assert_eq!(result.after_utf16_len(), 4);
        assert!(result.removed_object_references().is_empty());
        assert!(result.has_unknown_wire_fields());
    }

    #[test]
    fn cross_fragment_rewrite_preserves_wholly_untouched_fragment_records() {
        let first = raw_length_delimited(3, b"AA");
        let middle = raw_length_delimited(3, b"BB");
        let last = raw_length_delimited(3, b"CC");
        let unknown = raw_varint(99, 42);
        let source = [
            first.as_slice(),
            unknown.as_slice(),
            middle.as_slice(),
            last.as_slice(),
        ]
        .concat();

        let result = rewrite(&source, 2..4, "東京");
        let expected = [
            first.as_slice(),
            unknown.as_slice(),
            raw_length_delimited(3, "東京".as_bytes()).as_slice(),
            last.as_slice(),
        ]
        .concat();

        assert!(result.changed());
        assert_eq!(result.bytes(), expected);
        assert_eq!(decode(result.bytes()).text, ["AA", "東京", "CC"]);
    }

    #[test]
    fn insertion_at_fragment_boundary_has_right_affinity() {
        let first = raw_length_delimited(3, b"left");
        let second = raw_length_delimited(3, b"right");
        let source = [first.as_slice(), second.as_slice()].concat();

        let result = rewrite(&source, 4..4, "-");
        let expected = [
            first.as_slice(),
            raw_length_delimited(3, b"-right").as_slice(),
        ]
        .concat();

        assert_eq!(result.bytes(), expected);
    }

    #[test]
    fn astral_boundaries_are_checked_in_utf16_units() {
        let source = StorageArchive {
            text: vec!["A😀B".to_owned()],
            ..StorageArchive::default()
        }
        .encode_to_vec();

        assert_eq!(
            rewrite_storage_text_with_limits(&source, 2..3, "x", RewriteLimits::default()),
            Err(RewriteError::SurrogateSplit { index: 2 })
        );
        let result = rewrite(&source, 1..3, "🚀");
        assert_eq!(decode(result.bytes()).text, ["A🚀B"]);
        assert_eq!(result.before_utf16_len(), 4);
        assert_eq!(result.after_utf16_len(), 4);
    }

    #[test]
    fn generic_kernel_allows_format_owned_structural_controls() {
        let source = StorageArchive {
            text: vec!["Body".to_owned()],
            ..StorageArchive::default()
        }
        .encode_to_vec();

        let result = rewrite(&source, 4..4, "\u{4}\u{e}\u{fffc}");
        assert_eq!(decode(result.bytes()).text, ["Body\u{4}\u{e}\u{fffc}"]);
    }

    #[test]
    fn normalized_ranged_table_after_edit_stays_raw_exact() {
        let storage = StorageArchive {
            text: vec!["abcde".to_owned()],
            table_smartfield: Some(object_table(&[(0, None), (2, Some(20))])),
            ..StorageArchive::default()
        };
        let source = storage.encode_to_vec();
        let before_table = root_field(&source, 11);

        let result = rewrite(&source, 5..5, "!");

        assert_eq!(root_field(result.bytes(), 11), before_table);
        assert_eq!(
            decode(result.bytes()).table_smartfield,
            storage.table_smartfield
        );
    }

    #[test]
    fn prost_oracle_confirms_every_positional_table_policy() {
        let retained = object_table(&[(1, Some(101)), (3, Some(102)), (4, Some(103))]);
        let dropped = object_table(&[
            (1, Some(201)),
            (2, Some(202)),
            (3, Some(203)),
            (4, Some(204)),
        ]);
        let normalized = object_table(&[(0, None), (1, Some(301)), (4, Some(302))]);
        let storage = StorageArchive {
            text: vec!["abcde".to_owned()],
            table_para_style: Some(retained.clone()),
            table_para_data: Some(para_table(&[1, 3, 4])),
            table_list_style: Some(retained.clone()),
            table_char_style: Some(retained.clone()),
            table_attachment: Some(object_table(&[(1, Some(401)), (2, Some(402))])),
            table_smartfield: Some(normalized.clone()),
            table_layout_style: Some(retained.clone()),
            table_para_starts: Some(para_table(&[1, 3, 4])),
            table_bookmark: Some(normalized.clone()),
            table_footnote: Some(object_table(&[(1, Some(501)), (2, Some(502))])),
            table_section: Some(retained.clone()),
            table_rubyfield: Some(dropped.clone()),
            table_language: Some(string_table(&[1, 3, 4])),
            table_dictation: Some(string_table(&[1, 3, 4])),
            table_insertion: Some(dropped.clone()),
            table_deletion: Some(dropped.clone()),
            table_highlight: Some(normalized),
            table_para_bidi: Some(para_table(&[1, 3, 4])),
            table_overlapping_highlight: Some(overlapping_table(601)),
            table_pencil_annotation: Some(overlapping_table(701)),
            table_tatechuyoko: Some(dropped),
            table_drop_cap_style: Some(retained),
            ..StorageArchive::default()
        };

        let result = rewrite(&storage.encode_to_vec(), 1..3, "Z");
        let oracle = decode(result.bytes());

        assert_eq!(oracle.text, ["aZde"]);
        for table in [
            &oracle.table_para_style,
            &oracle.table_list_style,
            &oracle.table_char_style,
            &oracle.table_layout_style,
            &oracle.table_section,
            &oracle.table_drop_cap_style,
        ] {
            assert_eq!(object_indices(table), [1, 2, 3]);
        }
        for table in [
            &oracle.table_para_data,
            &oracle.table_para_starts,
            &oracle.table_para_bidi,
        ] {
            assert_eq!(para_indices(table), [1, 2, 3]);
        }
        for table in [&oracle.table_language, &oracle.table_dictation] {
            assert_eq!(string_indices(table), [1, 2, 3]);
        }
        assert!(oracle.table_attachment.is_none());
        assert!(oracle.table_footnote.is_none());
        for table in [
            &oracle.table_rubyfield,
            &oracle.table_insertion,
            &oracle.table_deletion,
            &oracle.table_tatechuyoko,
        ] {
            assert_eq!(object_indices(table), [2, 3]);
        }
        for table in [
            &oracle.table_smartfield,
            &oracle.table_bookmark,
            &oracle.table_highlight,
        ] {
            assert_eq!(object_indices(table), [0, 3]);
            assert!(
                table
                    .as_ref()
                    .is_some_and(|table| table.entries[0].object.is_none())
            );
        }
        for table in [
            &oracle.table_overlapping_highlight,
            &oracle.table_pencil_annotation,
        ] {
            let table = table
                .as_ref()
                .unwrap_or_else(|| panic!("overlapping table should remain"));
            assert_eq!(
                table
                    .entries
                    .iter()
                    .map(|entry| (entry.range.location, entry.range.length))
                    .collect::<Vec<_>>(),
                [(0, 1), (2, 1)]
            );
        }
    }

    #[test]
    fn unrelated_edit_refuses_unnormalized_ranged_table() {
        let source = StorageArchive {
            text: vec!["abcde".to_owned()],
            table_smartfield: Some(object_table(&[(2, Some(20))])),
            ..StorageArchive::default()
        }
        .encode_to_vec();

        assert!(matches!(
            rewrite_storage_text_with_limits(&source, 5..5, "!", RewriteLimits::default()),
            Err(RewriteError::InvalidFormat(message))
                if message.contains("refusing global normalization")
        ));
    }

    #[test]
    fn equal_text_default_preserves_exact_bytes_and_intersecting_tables() {
        let source = StorageArchive {
            text: vec!["A".to_owned()],
            table_smartfield: Some(object_table(&[(0, Some(44))])),
            ..StorageArchive::default()
        }
        .encode_to_vec();

        let result = rewrite(&source, 0..1, "A");
        let oracle = decode(result.bytes());

        assert!(!result.changed());
        assert_eq!(result.bytes(), source);
        assert_eq!(oracle.text, ["A"]);
        assert!(oracle.table_smartfield.is_some());
        assert!(result.removed_object_references().is_empty());
    }

    #[test]
    fn replace_selection_behavior_removes_equal_text_intersections() {
        let source = StorageArchive {
            text: vec!["A".to_owned()],
            table_smartfield: Some(object_table(&[(0, Some(44))])),
            ..StorageArchive::default()
        }
        .encode_to_vec();

        let result = rewrite_storage_text_with_behavior_and_limits(
            &source,
            0..1,
            "A",
            RewriteBehavior::ReplaceSelection,
            RewriteLimits::default(),
        )
        .unwrap_or_else(|error| panic!("legacy selection should rewrite: {error}"));
        let oracle = decode(result.bytes());

        assert!(result.changed());
        assert_eq!(oracle.text, ["A"]);
        assert!(oracle.table_smartfield.is_none());
        assert_eq!(result.removed_object_references(), [44]);
        assert_eq!(
            result.removed_object_references_by_field(),
            [RemovedObjectReference {
                storage_field_number: 11,
                identifier: 44,
            }]
        );
    }

    #[test]
    fn empty_and_unknown_only_attachment_and_footnote_tables_remain_raw() {
        let unknown = raw_varint(99, 7);
        for (attachment, footnote) in [(Vec::new(), unknown.clone()), (unknown.clone(), Vec::new())]
        {
            let source = [
                raw_length_delimited(3, b"abc"),
                raw_length_delimited(9, &attachment),
                raw_length_delimited(16, &footnote),
            ]
            .concat();
            let attachment_raw = root_field(&source, 9);
            let footnote_raw = root_field(&source, 16);

            let result = rewrite(&source, 0..0, "X");

            assert!(result.changed());
            assert_eq!(root_field(result.bytes(), 9), attachment_raw);
            assert_eq!(root_field(result.bytes(), 16), footnote_raw);
        }
    }

    #[test]
    fn unknown_table_data_keeps_container_when_final_known_entry_is_removed() {
        let unknown = raw_varint(99, 7);
        let mut attachment = object_table(&[(1, Some(81))]).encode_to_vec();
        attachment.extend_from_slice(&unknown);
        let mut footnote = object_table(&[(1, Some(82))]).encode_to_vec();
        footnote.extend_from_slice(&unknown);
        let source = [
            raw_length_delimited(3, b"abc"),
            raw_length_delimited(9, &attachment),
            raw_length_delimited(16, &footnote),
        ]
        .concat();

        let result = rewrite(&source, 1..2, "");

        assert_eq!(
            root_field(result.bytes(), 9),
            raw_length_delimited(9, &unknown)
        );
        assert_eq!(
            root_field(result.bytes(), 16),
            raw_length_delimited(16, &unknown)
        );
        assert_eq!(result.removed_object_references(), [81, 82]);
    }

    #[test]
    fn removed_reference_provenance_is_field_exact_and_aggregate_is_conservative() {
        let source = StorageArchive {
            text: vec!["abc".to_owned()],
            style_sheet: Some(reference(1)),
            table_char_style: Some(object_table(&[(0, Some(2))])),
            table_attachment: Some(object_table(&[(1, Some(2)), (2, Some(3))])),
            ..StorageArchive::default()
        }
        .encode_to_vec();

        let result = rewrite(&source, 1..3, "");

        assert_eq!(result.object_references_before(), [1, 2, 3]);
        assert_eq!(result.object_references_after(), [1, 2]);
        assert_eq!(result.object_reference_occurrences_before(), [1, 2, 2, 3]);
        assert_eq!(result.object_reference_occurrences_after(), [1, 2]);
        assert_eq!(result.removed_object_references(), [3]);
        assert_eq!(
            result.removed_object_references_by_field(),
            [
                RemovedObjectReference {
                    storage_field_number: 9,
                    identifier: 2,
                },
                RemovedObjectReference {
                    storage_field_number: 9,
                    identifier: 3,
                },
            ]
        );
    }

    #[test]
    fn noncanonical_untouched_framing_is_accepted_and_preserved() {
        let source = [0x9a, 0x80, 0x00, 0x81, 0x00, b'x'];

        let result = rewrite(&source, 0..1, "x");

        assert_eq!(result.bytes(), source);
        assert!(!result.changed());
    }

    #[test]
    fn nested_unknown_reference_field_is_reported_without_losing_raw_bytes() {
        let mut reference = reference(17).encode_to_vec();
        reference.extend_from_slice(&raw_varint(99, 5));
        let source = [
            raw_length_delimited(2, &reference),
            raw_length_delimited(3, b"x"),
        ]
        .concat();

        let result = rewrite(&source, 0..1, "x");

        assert!(result.has_unknown_wire_fields());
        assert_eq!(result.bytes(), source);
        assert_eq!(result.object_reference_occurrences_before(), [17]);
    }

    #[test]
    fn changed_table_preserves_unknown_root_table_entry_and_reference_bytes() {
        let reference_unknown = raw_varint(90, 77);
        let mut reference_payload = raw_varint(1, 17);
        reference_payload.extend_from_slice(&reference_unknown);
        let entry_unknown = raw_length_delimited(91, b"opaque-entry");
        let mut entry = raw_varint(1, 4);
        entry.extend_from_slice(&raw_length_delimited(2, &reference_payload));
        entry.extend_from_slice(&entry_unknown);
        let table_unknown = raw_varint(92, 88);
        let mut table = raw_length_delimited(1, &entry);
        table.extend_from_slice(&table_unknown);
        let root_unknown = raw_length_delimited(93, b"opaque-root");
        let source = [
            raw_length_delimited(3, b"abcde"),
            raw_length_delimited(8, &table),
            root_unknown.clone(),
        ]
        .concat();

        let result = rewrite(&source, 0..0, "X");
        let oracle = decode(result.bytes());

        assert!(result.has_unknown_wire_fields());
        assert_eq!(object_indices(&oracle.table_char_style), [5]);
        for raw in [
            reference_unknown.as_slice(),
            entry_unknown.as_slice(),
            table_unknown.as_slice(),
            root_unknown.as_slice(),
        ] {
            assert!(
                result
                    .bytes()
                    .windows(raw.len())
                    .any(|window| window == raw),
                "unknown raw field should remain byte-identical"
            );
        }
    }

    #[test]
    fn malformed_cardinality_order_wire_and_required_children_fail_closed() {
        let duplicate_root = [
            raw_length_delimited(3, b"x"),
            raw_length_delimited(5, &[]),
            raw_length_delimited(5, &[]),
        ]
        .concat();
        let wrong_table_wire = [raw_length_delimited(3, b"x"), raw_varint(5, 0)].concat();
        let unsorted_table = StorageArchive {
            text: vec!["abc".to_owned()],
            table_char_style: Some(object_table(&[(2, Some(1)), (1, Some(2))])),
            ..StorageArchive::default()
        }
        .encode_to_vec();
        let duplicate_index_entry = [raw_varint(1, 0), raw_varint(1, 1)].concat();
        let duplicate_index = [
            raw_length_delimited(3, b"x"),
            raw_length_delimited(8, &raw_length_delimited(1, &duplicate_index_entry)),
        ]
        .concat();
        let missing_reference_identifier =
            [raw_varint(1, 0), raw_length_delimited(2, &[])].concat();
        let missing_reference = [
            raw_length_delimited(3, b"x"),
            raw_length_delimited(8, &raw_length_delimited(1, &missing_reference_identifier)),
        ]
        .concat();
        let wrong_index_wire_entry = raw_length_delimited(1, b"zero");
        let wrong_index_wire = [
            raw_length_delimited(3, b"x"),
            raw_length_delimited(8, &raw_length_delimited(1, &wrong_index_wire_entry)),
        ]
        .concat();

        for malformed in [
            duplicate_root,
            wrong_table_wire,
            unsorted_table,
            duplicate_index,
            missing_reference,
            wrong_index_wire,
        ] {
            assert!(matches!(
                rewrite_storage_text_with_limits(&malformed, 0..0, "", RewriteLimits::default()),
                Err(RewriteError::InvalidFormat(_))
            ));
        }
    }

    #[test]
    fn malformed_overlapping_range_and_reference_are_rejected() {
        let range_without_length = raw_varint(1, 0);
        let entry_without_range_length = [
            raw_length_delimited(1, &range_without_length),
            raw_length_delimited(2, &reference(1).encode_to_vec()),
        ]
        .concat();
        let entry_without_reference = raw_length_delimited(
            1,
            &NativeRange {
                location: 0,
                length: 1,
            }
            .encode_to_vec(),
        );
        for entry in [entry_without_range_length, entry_without_reference] {
            let source = [
                raw_length_delimited(3, b"x"),
                raw_length_delimited(25, &raw_length_delimited(1, &entry)),
            ]
            .concat();
            assert!(matches!(
                rewrite_storage_text_with_limits(&source, 0..0, "", RewriteLimits::default()),
                Err(RewriteError::InvalidFormat(_))
            ));
        }
    }

    #[test]
    fn every_retained_resource_limit_is_enforced() {
        let source = [
            raw_length_delimited(3, b"ab"),
            raw_length_delimited(3, b"cd"),
        ]
        .concat();
        let message_limited =
            RewriteLimits::new(source.len() - 1, 64, 4, 4, 64, 16, 16, 1_024, 16_384)
                .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));
        assert!(matches!(
            rewrite_storage_text_with_limits(&source, 0..0, "", message_limited),
            Err(RewriteError::LimitExceeded {
                resource: "message bytes",
                ..
            })
        ));
        let field_limited = RewriteLimits::new(1_024, 1, 4, 4, 64, 16, 16, 1_024, 16_384)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));
        assert!(matches!(
            rewrite_storage_text_with_limits(&source, 0..0, "", field_limited),
            Err(RewriteError::LimitExceeded {
                resource: "fields",
                ..
            })
        ));
        let invalid_second_fragment = [
            raw_length_delimited(3, b"a"),
            raw_length_delimited(3, &[0xff]),
        ]
        .concat();
        assert!(matches!(
            rewrite_storage_text_with_limits(&invalid_second_fragment, 0..0, "", field_limited),
            Err(RewriteError::LimitExceeded {
                resource: "fields",
                ..
            })
        ));
        let fragment_limited = RewriteLimits::new(1_024, 64, 4, 1, 64, 16, 16, 1_024, 16_384)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));
        assert!(matches!(
            rewrite_storage_text_with_limits(&source, 0..0, "", fragment_limited),
            Err(RewriteError::LimitExceeded {
                resource: "text fragments",
                ..
            })
        ));

        let text_limited = RewriteLimits::new(1_024, 64, 4, 4, 3, 16, 16, 1_024, 16_384)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));
        assert!(matches!(
            rewrite_storage_text_with_limits(&source, 0..0, "", text_limited),
            Err(RewriteError::LimitExceeded {
                resource: "text bytes",
                ..
            })
        ));

        let table_source = StorageArchive {
            text: vec!["ab".to_owned()],
            table_char_style: Some(object_table(&[(0, Some(1)), (1, Some(2))])),
            ..StorageArchive::default()
        }
        .encode_to_vec();
        let table_limited = RewriteLimits::new(1_024, 64, 4, 4, 64, 1, 16, 1_024, 16_384)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));
        assert!(matches!(
            rewrite_storage_text_with_limits(&table_source, 0..0, "", table_limited),
            Err(RewriteError::LimitExceeded {
                resource: "table entries",
                ..
            })
        ));

        let reference_limited = RewriteLimits::new(1_024, 64, 4, 4, 64, 16, 1, 1_024, 16_384)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));
        assert!(matches!(
            rewrite_storage_text_with_limits(&table_source, 0..0, "", reference_limited),
            Err(RewriteError::LimitExceeded {
                resource: "object references",
                ..
            })
        ));

        let output_limited = RewriteLimits::new(1_024, 64, 4, 4, 64, 16, 16, 8, 16_384)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));
        assert!(matches!(
            rewrite_storage_text_with_limits(&source, 0..0, "much-longer", output_limited),
            Err(RewriteError::LimitExceeded {
                resource: "output bytes",
                ..
            })
        ));

        let work_limited = RewriteLimits::new(1_024, 64, 4, 4, 64, 16, 16, 1_024, 8)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));
        assert!(matches!(
            rewrite_storage_text_with_limits(&source, 0..0, "", work_limited),
            Err(RewriteError::LimitExceeded { .. })
        ));
        assert!(matches!(
            RewriteLimits::new(1_024, 64, 3, 4, 64, 16, 16, 1_024, 16_384),
            Err(RewriteError::InvalidLimit {
                field: "nesting",
                ..
            })
        ));
    }

    #[test]
    fn prepared_execution_max_minus_one_refuses_before_any_execution_allocation() {
        let source = StorageArchive {
            text: vec!["abc".to_owned()],
            table_char_style: Some(object_table(&[(0, Some(11)), (2, Some(12))])),
            ..StorageArchive::default()
        }
        .encode_to_vec();
        let prepare = || {
            prepare_storage_text_rewrite_with_behavior_and_limits(
                &source,
                1..2,
                "longer",
                RewriteBehavior::PreserveOnEqualText,
                RewriteLimits::default(),
            )
            .unwrap_or_else(|error| panic!("storage plan should prepare: {error}"))
        };
        let requirements = prepare().execution_requirements();
        assert!(requirements.output_bytes() > 0);
        assert!(requirements.retained_elements() > 0);
        assert!(requirements.retained_bytes() > 0);
        assert!(requirements.peak_scratch_bytes() > 0);
        assert!(requirements.allocations() > 0);
        assert!(requirements.work() > 0);

        for axis in 0..6 {
            STORAGE_EXECUTION_ENTRIES.with(|entries| entries.set(0));
            let mut limits = StorageRewriteExecutionLimits {
                max_output_bytes: requirements.output_bytes(),
                max_retained_elements: requirements.retained_elements(),
                max_retained_bytes: requirements.retained_bytes(),
                max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
                max_allocations: requirements.allocations(),
                max_work: requirements.work(),
            };
            match axis {
                0 => limits.max_output_bytes -= 1,
                1 => limits.max_retained_elements -= 1,
                2 => limits.max_retained_bytes -= 1,
                3 => limits.max_peak_scratch_bytes -= 1,
                4 => limits.max_allocations -= 1,
                5 => limits.max_work -= 1,
                _ => unreachable!(),
            }
            assert!(matches!(
                prepare().execute(limits),
                Err(RewriteError::LimitExceeded { .. })
            ));
            STORAGE_EXECUTION_ENTRIES.with(|entries| assert_eq!(entries.get(), 0));
        }

        STORAGE_EXECUTION_ENTRIES.with(|entries| entries.set(0));
        let output = prepare()
            .execute(StorageRewriteExecutionLimits {
                max_output_bytes: requirements.output_bytes(),
                max_retained_elements: requirements.retained_elements(),
                max_retained_bytes: requirements.retained_bytes(),
                max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
                max_allocations: requirements.allocations(),
                max_work: requirements.work(),
            })
            .unwrap_or_else(|error| panic!("exact execution should succeed: {error}"));
        assert!(output.changed());
        STORAGE_EXECUTION_ENTRIES.with(|entries| assert_eq!(entries.get(), 1));
    }

    #[test]
    fn validation_only_api_checks_full_tree_without_an_output_budget() {
        let mut nested_reference = reference(9).encode_to_vec();
        nested_reference.extend_from_slice(&raw_varint(99, 1));
        let entry = [raw_varint(1, 1), raw_length_delimited(2, &nested_reference)].concat();
        let source = [
            raw_length_delimited(2, &reference(7).encode_to_vec()),
            raw_length_delimited(3, b"A"),
            raw_length_delimited(3, "😀".as_bytes()),
            raw_length_delimited(8, &raw_length_delimited(1, &entry)),
        ]
        .concat();
        let limits = RewriteLimits::new(1_024, 64, 4, 4, 64, 8, 8, 1, 16_384)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));

        let validation = validate_storage_with_limits(&source, limits)
            .unwrap_or_else(|error| panic!("full storage should validate: {error}"));

        assert_eq!(validation.utf8_len(), 5);
        assert_eq!(validation.utf16_len(), 3);
        assert_eq!(validation.fragments(), 2);
        assert_eq!(validation.fields(), 10);
        assert_eq!(validation.table_entries(), 1);
        assert_eq!(validation.reference_occurrences(), 2);
        assert!(validation.has_unknown_wire_fields());

        let missing_identifier_entry = [raw_varint(1, 0), raw_length_delimited(2, &[])].concat();
        let malformed = [
            raw_length_delimited(3, b"x"),
            raw_length_delimited(8, &raw_length_delimited(1, &missing_identifier_entry)),
        ]
        .concat();
        assert!(matches!(
            validate_storage_with_limits(&malformed, RewriteLimits::default()),
            Err(RewriteError::InvalidFormat(_))
        ));
    }

    #[test]
    fn reversed_oob_and_fragment_wire_errors_are_typed() {
        let source = raw_length_delimited(3, b"abc");
        assert!(matches!(
            rewrite_storage_text_with_limits(
                &source,
                Range { start: 2, end: 1 },
                "",
                RewriteLimits::default()
            ),
            Err(RewriteError::ReversedRange { start: 2, end: 1 })
        ));
        assert!(matches!(
            rewrite_storage_text_with_limits(&source, 0..4, "", RewriteLimits::default()),
            Err(RewriteError::RangeOutOfBounds {
                index: 4,
                length: 3
            })
        ));
        assert!(matches!(
            rewrite_storage_text_with_limits(
                &[0x18, 0x01],
                0..0,
                "",
                RewriteLimits::default()
            ),
            Err(RewriteError::InvalidFormat(message)) if message.contains("wire type")
        ));
    }
}
