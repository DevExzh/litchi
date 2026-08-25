//! Strict generated-free Pages body-footnote graph codec.
//!
//! This module owns the small source-authoritative wire seam needed by a
//! focused body-footnote owner.  It can create the four canonical payloads
//! for a new footnote graph and can replace the repeated body-table entries
//! without reconstructing any pre-existing protobuf message.  Existing entry
//! and root unknown fields, including balanced groups and overlong unknown
//! scalar values, are copied byte-for-byte.  Known fields are always emitted
//! with canonical keys, lengths, and scalar values.
//!
//! The generated Buffa projections remain private to their sibling codecs.
//! This module only calls their generated-free APIs for candidate validation;
//! no generated type is present in its public surface.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict wire pass intentionally precedes low-level helpers."
)]

use core::fmt;
use std::{mem::size_of, num::NonZeroU64, str};

use crate::{
    pages_body_codec, pages_footnote_codec, pages_footnote_marker_codec, text_storage_codec,
};

const FOOTNOTE_REFERENCE_SUPER_FIELD: u32 = 1;
const FOOTNOTE_REFERENCE_STORAGE_FIELD: u32 = 2;
const FOOTNOTE_REFERENCE_CUSTOM_MARK_FIELD: u32 = 3;
const TEXTUAL_KIND_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const STORAGE_KIND_FIELD: u32 = 1;
const STORAGE_STYLESHEET_FIELD: u32 = 2;
const STORAGE_TEXT_FIELD: u32 = 3;
const STORAGE_PARA_STYLE_FIELD: u32 = 5;
const STORAGE_PARA_DATA_FIELD: u32 = 6;
const STORAGE_LIST_STYLE_FIELD: u32 = 7;
const STORAGE_ATTACHMENT_TABLE_FIELD: u32 = 9;
const STORAGE_IN_DOCUMENT_FIELD: u32 = 10;
const STORAGE_PARA_STARTS_FIELD: u32 = 14;
const STORAGE_LANGUAGE_FIELD: u32 = 19;
const STORAGE_PARA_BIDI_FIELD: u32 = 24;
const STORAGE_DROP_CAP_FIELD: u32 = 28;
const TABLE_ENTRIES_FIELD: u32 = 1;
const BOUNDARY_CHARACTER_INDEX_FIELD: u32 = 1;
const BOUNDARY_OBJECT_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const FOOTNOTE_KIND: u64 = 2;
const FOOTNOTE_MARK_KIND: i32 = 2;
const FOOTNOTE_TEXT_PREFIX: &str = "\u{fffc} ";
const MAX_RECURSION: u32 = 64;

/// Finite policy shared by graph creation, table rewrites, and candidate
/// readback.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_nesting: u32,
    max_entries: usize,
}

impl DecodeOptions {
    /// Construct an explicit finite input/output, field, work, nesting, and
    /// body-entry policy.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_nesting: u32,
        max_entries: usize,
    ) -> Self {
        Self {
            max_input_bytes,
            max_output_bytes,
            max_fields,
            max_work_bytes,
            max_nesting,
            max_entries,
        }
    }

    /// Build a conservative profile from one source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let input = source.len().max(1);
        Self::new(
            input,
            input.saturating_mul(2).max(1),
            input.saturating_mul(8).max(1),
            input.saturating_mul(32).max(1),
            8,
            input.max(1),
        )
    }

    /// Replace the candidate output ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the source/candidate input ceiling used by readback.
    #[must_use]
    pub const fn with_max_input_bytes(mut self, maximum: usize) -> Self {
        self.max_input_bytes = maximum;
        self
    }

    /// Replace the maximum number of body-footnote entries.
    #[must_use]
    pub const fn with_max_entries(mut self, maximum: usize) -> Self {
        self.max_entries = maximum;
        self
    }

    /// Maximum input bytes accepted by this policy.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Maximum candidate bytes accepted by this policy.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Maximum strict field visits accepted by this policy.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Maximum aggregate traversal work accepted by this policy.
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    /// Maximum protobuf/group nesting accepted by this policy.
    #[must_use]
    pub const fn max_nesting(self) -> u32 {
        self.max_nesting
    }

    /// Maximum body-table entry count accepted by this policy.
    #[must_use]
    pub const fn max_entries(self) -> usize {
        self.max_entries
    }
}

/// Typed finite resource failure from the graph codec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Source payload exceeded its input ceiling.
    InputBytes { observed: usize, maximum: usize },
    /// Candidate payload exceeded its output ceiling.
    OutputBytes { observed: usize, maximum: usize },
    /// Strict field visits exceeded their ceiling.
    Fields { observed: usize, maximum: usize },
    /// Aggregate traversal/rewrite work exceeded its ceiling.
    WorkBytes { observed: usize, maximum: usize },
    /// Group or message nesting exceeded its ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// Body-table entry count exceeded its ceiling.
    Entries { observed: usize, maximum: usize },
    /// Text bytes exceeded the finite graph-creation ceiling.
    TextBytes { observed: usize, maximum: usize },
}

/// Strict graph codec failure.  Diagnostics contain schema labels and finite
/// observations only; authored text and raw bytes are never formatted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(&'static str),
    Limit(DecodeLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    DuplicateKey(&'static str),
    InvalidValue(&'static str),
    InvalidOrdering,
    InvalidCandidate,
    Allocation { amount: usize },
}

impl DecodeError {
    const fn wire(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Wire(reason),
        }
    }

    const fn limit(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Limit(limit),
        }
    }

    const fn missing(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::MissingRequired(field),
        }
    }

    const fn duplicate(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateSingular(field),
        }
    }

    const fn duplicate_key(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateKey(field),
        }
    }

    const fn invalid(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::InvalidValue(reason),
        }
    }

    const fn ordering() -> Self {
        Self {
            kind: DecodeErrorKind::InvalidOrdering,
        }
    }

    const fn candidate() -> Self {
        Self {
            kind: DecodeErrorKind::InvalidCandidate,
        }
    }

    const fn allocation(amount: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation { amount },
        }
    }

    /// Return the finite resource observation, when applicable.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Limit(limit) => Some(limit),
            _ => None,
        }
    }

    /// Return the failed allocation amount, when applicable.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        match self.kind {
            DecodeErrorKind::Allocation { amount } => Some(amount),
            _ => None,
        }
    }

    /// Return the missing required field, when applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.kind {
            DecodeErrorKind::Wire(reason) => {
                write!(formatter, "Pages footnote graph wire error: {reason}")
            },
            DecodeErrorKind::Limit(DecodeLimit::InputBytes { observed, maximum }) => write!(
                formatter,
                "Pages footnote graph input is {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::OutputBytes { observed, maximum }) => write!(
                formatter,
                "Pages footnote graph output is {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "Pages footnote graph visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::WorkBytes { observed, maximum }) => write!(
                formatter,
                "Pages footnote graph requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Pages footnote graph nesting is {observed}; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Entries { observed, maximum }) => write!(
                formatter,
                "Pages footnote graph has {observed} entries; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::TextBytes { observed, maximum }) => write!(
                formatter,
                "Pages footnote graph text is {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => write!(formatter, "duplicate {field}"),
            DecodeErrorKind::DuplicateKey(field) => write!(formatter, "duplicate {field}"),
            DecodeErrorKind::InvalidValue(reason) => {
                write!(formatter, "invalid Pages footnote graph value: {reason}")
            },
            DecodeErrorKind::InvalidOrdering => {
                formatter.write_str("Pages body-footnote entries are not strictly ordered")
            },
            DecodeErrorKind::InvalidCandidate => {
                formatter.write_str("Pages footnote graph candidate failed verification")
            },
            DecodeErrorKind::Allocation { amount } => write!(
                formatter,
                "cannot allocate {amount} Pages footnote graph bytes"
            ),
        }
    }
}

impl std::error::Error for DecodeError {}

/// Exact accounting for one strict body-table decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    entries: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl DecodeReport {
    /// Source payload bytes inspected.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Alias for source-byte terminology.
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.input_bytes
    }

    /// Strict field visits, including unknown groups.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate bounded traversal work.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum observed group/message depth.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Body-table entry count.
    #[must_use]
    pub const fn entries(self) -> usize {
        self.entries
    }

    /// Bytes retained by the returned borrowed snapshot.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Temporary bytes used by the strict staging pass.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Conservative result resources computed before a candidate output buffer is
/// reserved.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeResourceUpperBound {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    entries: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl DecodeResourceUpperBound {
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.input_bytes
    }

    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    #[must_use]
    pub const fn entries(self) -> usize {
        self.entries
    }

    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Exact accounting for a body-table candidate rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BodyFootnoteTableRewriteReport {
    source: DecodeReport,
    result: DecodeReport,
    output_bytes: usize,
    rewrite_work_bytes: usize,
    entries_before: usize,
    entries_after: usize,
    inserted: usize,
    removed: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    allocations: usize,
}

impl BodyFootnoteTableRewriteReport {
    /// Source decode accounting.
    #[must_use]
    pub const fn source(self) -> DecodeReport {
        self.source
    }

    /// Candidate decode accounting.
    #[must_use]
    pub const fn result(self) -> DecodeReport {
        self.result
    }

    /// Exact candidate payload size.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Work charged to size, emit, and verify the rewrite.
    #[must_use]
    pub const fn rewrite_work_bytes(self) -> usize {
        self.rewrite_work_bytes
    }

    /// Source entry count.
    #[must_use]
    pub const fn entries_before(self) -> usize {
        self.entries_before
    }

    /// Candidate entry count.
    #[must_use]
    pub const fn entries_after(self) -> usize {
        self.entries_after
    }

    /// Number of newly encoded entries.
    #[must_use]
    pub const fn inserted(self) -> usize {
        self.inserted
    }

    /// Number of source entries omitted by the candidate.
    #[must_use]
    pub const fn removed(self) -> usize {
        self.removed
    }

    /// Candidate bytes retained by the returned vector.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Temporary staging bytes accounted by the rewrite.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Number of output-buffer allocations performed.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

/// Borrowed body-table entry facts.  `raw` is the complete nested
/// `ObjectAttribute` payload; `field_raw` includes its repeated field-1 key
/// and length prefix.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BodyFootnoteEntrySnapshot<'source> {
    raw: &'source [u8],
    field_raw: &'source [u8],
    character_index: u32,
    reference_identifier: NonZeroU64,
}

impl<'source> BodyFootnoteEntrySnapshot<'source> {
    /// UTF-16 body position carried by this entry.
    #[must_use]
    pub const fn character_index(self) -> u32 {
        self.character_index
    }

    /// Native reference identity carried by this entry.
    #[must_use]
    pub const fn reference_identifier(self) -> NonZeroU64 {
        self.reference_identifier
    }

    /// Exact nested entry bytes.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }

    /// Exact repeated field bytes, including its canonical outer framing.
    #[must_use]
    pub const fn field_raw(self) -> &'source [u8] {
        self.field_raw
    }
}

/// Borrowed table view with allocation-free source-order entry iteration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BodyFootnoteTableSnapshot<'source> {
    source: &'source [u8],
    entries: usize,
}

impl<'source> BodyFootnoteTableSnapshot<'source> {
    /// Exact source table bytes.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }

    /// Number of repeated body-footnote entries.
    #[must_use]
    pub const fn len(self) -> usize {
        self.entries
    }

    /// Whether the table has no entries.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.entries == 0
    }

    /// Iterate complete borrowed entries in source order.
    #[must_use]
    pub fn entries(self) -> BodyFootnoteEntryIter<'source> {
        BodyFootnoteEntryIter {
            source: self.source,
            yielded: 0,
            maximum: self.entries,
        }
    }
}

/// Allocation-free body-table entry iterator.
#[derive(Debug, Clone, Copy)]
pub struct BodyFootnoteEntryIter<'source> {
    source: &'source [u8],
    yielded: usize,
    maximum: usize,
}

impl<'source> Iterator for BodyFootnoteEntryIter<'source> {
    type Item = BodyFootnoteEntrySnapshot<'source>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.yielded >= self.maximum {
            return None;
        }
        while !self.source.is_empty() {
            let field = parse_field_unmetered(&mut self.source).ok()??;
            if field.number != TABLE_ENTRIES_FIELD {
                continue;
            }
            let payload = field.payload?;
            let entry = parse_entry_unmetered(payload).ok()?;
            self.yielded = self.yielded.saturating_add(1);
            return Some(BodyFootnoteEntrySnapshot {
                raw: payload,
                field_raw: field.raw,
                character_index: entry.character_index,
                reference_identifier: entry.reference_identifier,
            });
        }
        None
    }
}

/// One requested output entry.  `preserve` copies both nested and enclosing
/// source bytes exactly; `new` emits a canonical entry with no unknowns.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BodyFootnoteEntryWrite<'source> {
    raw: Option<&'source [u8]>,
    field_raw: Option<&'source [u8]>,
    character_index: u32,
    reference_identifier: NonZeroU64,
}

impl<'source> BodyFootnoteEntryWrite<'source> {
    /// Build a canonical new body-footnote entry.
    #[must_use]
    pub const fn new(character_index: u32, reference_identifier: NonZeroU64) -> Self {
        Self {
            raw: None,
            field_raw: None,
            character_index,
            reference_identifier,
        }
    }

    /// Retain an existing entry's complete source framing and unknown fields.
    #[must_use]
    pub const fn preserve(entry: BodyFootnoteEntrySnapshot<'source>) -> Self {
        Self {
            raw: Some(entry.raw),
            field_raw: Some(entry.field_raw),
            character_index: entry.character_index,
            reference_identifier: entry.reference_identifier,
        }
    }

    /// Entry UTF-16 position.
    #[must_use]
    pub const fn character_index(self) -> u32 {
        self.character_index
    }

    /// Entry reference identity.
    #[must_use]
    pub const fn reference_identifier(self) -> NonZeroU64 {
        self.reference_identifier
    }

    /// Whether this request retains source bytes rather than creating a
    /// canonical nested entry.
    #[must_use]
    pub const fn preserves_raw(self) -> bool {
        self.raw.is_some() && self.field_raw.is_some()
    }
}

/// Ordered body-table rewrite request.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BodyFootnoteTableWrite<'entries, 'source> {
    entries: &'entries [BodyFootnoteEntryWrite<'source>],
}

impl<'entries, 'source> BodyFootnoteTableWrite<'entries, 'source> {
    /// Build an ordered entry request.
    #[must_use]
    pub const fn new(entries: &'entries [BodyFootnoteEntryWrite<'source>]) -> Self {
        Self { entries }
    }

    /// Borrow requested entries in their output order.
    #[must_use]
    pub const fn entries(self) -> &'entries [BodyFootnoteEntryWrite<'source>] {
        self.entries
    }
}

/// Graph creation input.  IDs are checked nonzero before this type can be
/// constructed by callers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FootnoteGraphWrite<'text> {
    reference_identifier: NonZeroU64,
    storage_identifier: NonZeroU64,
    marker_identifier: NonZeroU64,
    character_index: u32,
    text: &'text str,
    custom_mark: Option<&'text str>,
    stylesheet_identifier: Option<NonZeroU64>,
    paragraph_style_identifier: Option<NonZeroU64>,
    list_style_identifier: Option<NonZeroU64>,
    language: Option<&'text str>,
}

impl<'text> FootnoteGraphWrite<'text> {
    /// Build a canonical graph creation request.
    #[must_use]
    pub const fn new(
        reference_identifier: NonZeroU64,
        storage_identifier: NonZeroU64,
        marker_identifier: NonZeroU64,
        character_index: u32,
        text: &'text str,
    ) -> Self {
        Self {
            reference_identifier,
            storage_identifier,
            marker_identifier,
            character_index,
            text,
            custom_mark: None,
            stylesheet_identifier: None,
            paragraph_style_identifier: None,
            list_style_identifier: None,
            language: None,
        }
    }

    /// Set the optional custom marker field, retaining explicit empty-string
    /// presence when `Some("")` is supplied.
    #[must_use]
    pub const fn with_custom_mark(mut self, custom_mark: Option<&'text str>) -> Self {
        self.custom_mark = custom_mark;
        self
    }

    /// Supply the body-storage style/language template used by native Pages
    /// footnote storage creation.  All supplied references are emitted as
    /// canonical index-zero table entries and remain outside the public API's
    /// raw-ID surface.
    #[must_use]
    pub const fn with_storage_template(
        mut self,
        stylesheet_identifier: Option<NonZeroU64>,
        paragraph_style_identifier: Option<NonZeroU64>,
        list_style_identifier: Option<NonZeroU64>,
        language: Option<&'text str>,
    ) -> Self {
        self.stylesheet_identifier = stylesheet_identifier;
        self.paragraph_style_identifier = paragraph_style_identifier;
        self.list_style_identifier = list_style_identifier;
        self.language = language;
        self
    }

    /// Reference object identity.
    #[must_use]
    pub const fn reference_identifier(self) -> NonZeroU64 {
        self.reference_identifier
    }

    /// Storage object identity.
    #[must_use]
    pub const fn storage_identifier(self) -> NonZeroU64 {
        self.storage_identifier
    }

    /// Marker object identity.
    #[must_use]
    pub const fn marker_identifier(self) -> NonZeroU64 {
        self.marker_identifier
    }

    /// UTF-16 position of the body anchor.
    #[must_use]
    pub const fn character_index(self) -> u32 {
        self.character_index
    }

    /// Borrow requested text.
    #[must_use]
    pub const fn text(self) -> &'text str {
        self.text
    }

    /// Borrow requested custom marker.
    #[must_use]
    pub const fn custom_mark(self) -> Option<&'text str> {
        self.custom_mark
    }

    /// Optional style-sheet object copied from the body template.
    #[must_use]
    pub const fn stylesheet_identifier(self) -> Option<NonZeroU64> {
        self.stylesheet_identifier
    }

    /// Optional paragraph-style object copied from the body template.
    #[must_use]
    pub const fn paragraph_style_identifier(self) -> Option<NonZeroU64> {
        self.paragraph_style_identifier
    }

    /// Optional list-style object copied from the body template.
    #[must_use]
    pub const fn list_style_identifier(self) -> Option<NonZeroU64> {
        self.list_style_identifier
    }

    /// Optional language marker copied from the body template.
    #[must_use]
    pub const fn language(self) -> Option<&'text str> {
        self.language
    }
}

/// Canonical creation payloads for one body-footnote graph.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FootnoteGraphPayloads {
    reference: Vec<u8>,
    storage: Vec<u8>,
    marker: Vec<u8>,
    body_entry: Vec<u8>,
}

impl FootnoteGraphPayloads {
    /// Reference-attachment message payload.
    #[must_use]
    pub fn reference(&self) -> &[u8] {
        &self.reference
    }

    /// Footnote storage message payload.
    #[must_use]
    pub fn storage(&self) -> &[u8] {
        &self.storage
    }

    /// Textual marker message payload.
    #[must_use]
    pub fn marker(&self) -> &[u8] {
        &self.marker
    }

    /// Body-table entry payload including its repeated field-1 framing.
    #[must_use]
    pub fn body_entry(&self) -> &[u8] {
        &self.body_entry
    }
}

/// Exact accounting for canonical graph creation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct GraphEncodeReport {
    reference_bytes: usize,
    storage_bytes: usize,
    marker_bytes: usize,
    body_entry_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl GraphEncodeReport {
    /// Reference payload bytes.
    #[must_use]
    pub const fn reference_bytes(self) -> usize {
        self.reference_bytes
    }

    /// Storage payload bytes.
    #[must_use]
    pub const fn storage_bytes(self) -> usize {
        self.storage_bytes
    }

    /// Marker payload bytes.
    #[must_use]
    pub const fn marker_bytes(self) -> usize {
        self.marker_bytes
    }

    /// Body entry bytes including field framing.
    #[must_use]
    pub const fn body_entry_bytes(self) -> usize {
        self.body_entry_bytes
    }

    /// Aggregate output bytes.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Canonical fields emitted and candidate fields inspected.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate bounded creation and candidate work.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Number of payload allocations.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Bytes retained by returned payloads.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Temporary bytes used while sizing/validating.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Decode one body-footnote table with exact accounting.
pub fn decode_body_footnote_table(
    source: &[u8],
    options: DecodeOptions,
) -> Result<BodyFootnoteTableSnapshot<'_>, DecodeError> {
    Ok(decode_body_footnote_table_with_report(source, options)?.0)
}

/// Decode one body-footnote table with exact source/resource reporting.
pub fn decode_body_footnote_table_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(BodyFootnoteTableSnapshot<'_>, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let entries = scan_table(source, options, &mut budget)?;
    Ok((
        BodyFootnoteTableSnapshot { source, entries },
        budget.report(entries),
    ))
}

/// Compatibility alias using the shorter table spelling.
pub fn decode_footnote_table_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(BodyFootnoteTableSnapshot<'_>, DecodeReport), DecodeError> {
    decode_body_footnote_table_with_report(source, options)
}

/// Rewrite body-footnote table entries while retaining all untouched source
/// fields and nested entry bytes.
pub fn rewrite_body_footnote_table(
    source: &[u8],
    write: BodyFootnoteTableWrite<'_, '_>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    Ok(rewrite_body_footnote_table_with_report(source, write, options)?.0)
}

/// Rewrite body-footnote table entries with source/result reports and a
/// candidate strict readback performed before returning the output.
pub fn rewrite_body_footnote_table_with_report(
    source: &[u8],
    write: BodyFootnoteTableWrite<'_, '_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, BodyFootnoteTableRewriteReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let (fields, source_entries) = scan_table_fields(source, options, &mut budget)?;
    let source_report = budget.report(source_entries.len());
    validate_requested_entries(write.entries(), options, &mut budget)?;
    let output_bytes = measured_table_output(&fields, source_entries.len(), write.entries())?;
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::limit(DecodeLimit::OutputBytes {
            observed: output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    let rewrite_work_bytes = source
        .len()
        .checked_add(output_bytes)
        .and_then(|value| value.checked_add(source_entries.len().saturating_mul(16)))
        .and_then(|value| value.checked_add(write.entries().len().saturating_mul(16)))
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::WorkBytes {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    let total_work = budget
        .work_bytes
        .checked_add(rewrite_work_bytes)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::WorkBytes {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    if total_work > options.max_work_bytes {
        return Err(DecodeError::limit(DecodeLimit::WorkBytes {
            observed: total_work,
            maximum: options.max_work_bytes,
        }));
    }
    let mut output = reserve_output(output_bytes)?;
    emit_table(&fields, source_entries.len(), write.entries(), &mut output)?;
    if output.len() != output_bytes {
        return Err(DecodeError::candidate());
    }
    let candidate_options = options
        .with_max_output_bytes(options.max_output_bytes.max(output.len()))
        .with_max_entries(options.max_entries.max(write.entries().len()))
        .with_max_input_bytes(options.max_input_bytes.max(output.len()));
    let (candidate, result_report) =
        decode_body_footnote_table_with_report(&output, candidate_options)?;
    verify_candidate_entries(candidate, write.entries())?;
    let inserted = write
        .entries()
        .iter()
        .filter(|entry| !entry.preserves_raw())
        .count();
    let removed = source_entries.len().saturating_sub(write.entries().len());
    let scratch_bytes = fields
        .len()
        .checked_mul(size_of::<RawField<'_>>())
        .and_then(|value| {
            value.checked_add(
                source_entries
                    .len()
                    .saturating_mul(size_of::<ParsedEntry>()),
            )
        })
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::WorkBytes {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    Ok((
        output,
        BodyFootnoteTableRewriteReport {
            source: source_report,
            result: result_report,
            output_bytes,
            rewrite_work_bytes,
            entries_before: source_entries.len(),
            entries_after: write.entries().len(),
            inserted,
            removed,
            retained_bytes: output_bytes,
            scratch_bytes,
            allocations: 1,
        },
    ))
}

/// Compatibility alias using the shorter table spelling.
pub fn rewrite_footnote_table_with_report(
    source: &[u8],
    write: BodyFootnoteTableWrite<'_, '_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, BodyFootnoteTableRewriteReport), DecodeError> {
    rewrite_body_footnote_table_with_report(source, write, options)
}

/// Create canonical reference, storage, marker, and body-entry payloads for
/// one new body-footnote graph.
pub fn encode_footnote_graph(
    write: FootnoteGraphWrite<'_>,
    options: DecodeOptions,
) -> Result<FootnoteGraphPayloads, DecodeError> {
    Ok(encode_footnote_graph_with_report(write, options)?.0)
}

/// Create canonical graph payloads and return exact output/resource accounting.
pub fn encode_footnote_graph_with_report(
    write: FootnoteGraphWrite<'_>,
    options: DecodeOptions,
) -> Result<(FootnoteGraphPayloads, GraphEncodeReport), DecodeError> {
    validate_graph_write(write, options)?;
    let text_bytes = FOOTNOTE_TEXT_PREFIX
        .len()
        .checked_add(write.text().len())
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::OutputBytes {
                observed: usize::MAX,
                maximum: options.max_output_bytes,
            })
        })?;
    let reference_bytes =
        canonical_reference_payload_size(write.storage_identifier(), write.custom_mark())?;
    let storage_bytes = canonical_storage_payload_size(text_bytes, write)?;
    let marker_bytes = canonical_marker_payload_size();
    let body_entry_bytes =
        canonical_body_entry_payload_size(write.character_index(), write.reference_identifier())?;
    let output_bytes = reference_bytes
        .checked_add(storage_bytes)
        .and_then(|value| value.checked_add(marker_bytes))
        .and_then(|value| value.checked_add(body_entry_bytes))
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::OutputBytes {
                observed: usize::MAX,
                maximum: options.max_output_bytes,
            })
        })?;
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::limit(DecodeLimit::OutputBytes {
            observed: output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    let fields = canonical_graph_field_count(write);
    if fields > options.max_fields {
        return Err(DecodeError::limit(DecodeLimit::Fields {
            observed: fields,
            maximum: options.max_fields,
        }));
    }
    let work_bytes = output_bytes
        .checked_mul(2)
        .and_then(|value| value.checked_add(text_bytes))
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::WorkBytes {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    if work_bytes > options.max_work_bytes {
        return Err(DecodeError::limit(DecodeLimit::WorkBytes {
            observed: work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    let text = format_text(write.text)?;
    let marker = canonical_marker_payload();
    let storage = canonical_storage_payload(&text, write);
    let reference = canonical_reference_payload(write.storage_identifier(), write.custom_mark());
    let body_entry =
        canonical_body_entry_payload(write.character_index(), write.reference_identifier());
    let measured_output_bytes = reference
        .len()
        .checked_add(storage.len())
        .and_then(|value| value.checked_add(marker.len()))
        .and_then(|value| value.checked_add(body_entry.len()))
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::OutputBytes {
                observed: usize::MAX,
                maximum: options.max_output_bytes,
            })
        })?;
    if measured_output_bytes != output_bytes
        || text.len() != text_bytes
        || reference.len() != reference_bytes
        || storage.len() != storage_bytes
        || marker.len() != marker_bytes
        || body_entry.len() != body_entry_bytes
    {
        return Err(DecodeError::candidate());
    }
    verify_graph_payloads(&reference, &storage, &marker, &body_entry, write, options)?;
    Ok((
        FootnoteGraphPayloads {
            reference,
            storage,
            marker,
            body_entry,
        },
        GraphEncodeReport {
            reference_bytes,
            storage_bytes,
            marker_bytes,
            body_entry_bytes,
            output_bytes,
            fields,
            work_bytes,
            allocations: 4,
            retained_bytes: output_bytes,
            scratch_bytes: text.len(),
        },
    ))
}

#[derive(Debug, Clone, Copy)]
struct RawField<'source> {
    number: u32,
    wire: u8,
    raw: &'source [u8],
    payload: Option<&'source [u8]>,
    value: Option<u64>,
    canonical_value: bool,
}

#[derive(Debug, Clone, Copy)]
struct ParsedEntry {
    character_index: u32,
    reference_identifier: NonZeroU64,
}

#[derive(Debug, Clone, Copy)]
enum UnmeteredValue<'source> {
    Varint(u64),
    Fixed64,
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32,
}

#[derive(Debug, Clone, Copy)]
struct UnmeteredField<'source> {
    number: u32,
    raw: &'source [u8],
    payload: Option<&'source [u8]>,
    value: UnmeteredValue<'source>,
}

#[derive(Debug)]
struct Budget {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_nesting: u32,
    max_depth: u32,
}

impl Budget {
    const fn new(source: &[u8], options: DecodeOptions) -> Self {
        Self {
            input_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
            max_nesting: options.max_nesting,
            max_depth: 0,
        }
    }

    fn charge_field(&mut self) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(1);
        if observed > self.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed,
                maximum: self.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self.work_bytes.checked_add(amount).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::WorkBytes {
                observed: usize::MAX,
                maximum: self.max_work_bytes,
            })
        })?;
        if observed > self.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::WorkBytes {
                observed,
                maximum: self.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn charge_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.max_nesting {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.max_nesting,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    const fn report(&self, entries: usize) -> DecodeReport {
        DecodeReport {
            input_bytes: self.input_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            entries,
            retained_bytes: self.input_bytes,
            scratch_bytes: 0,
        }
    }
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    // Sibling Buffa codecs apply their own hard message ceiling during the
    // candidate checks.  This orchestration layer only owns its configured
    // finite source/output ceilings and therefore does not depend on Buffa's
    // generated/runtime constants.
    let maximum = usize::MAX;
    if options.max_input_bytes > maximum {
        return Err(DecodeError::limit(DecodeLimit::InputBytes {
            observed: options.max_input_bytes,
            maximum,
        }));
    }
    if source.len() > options.max_input_bytes {
        return Err(DecodeError::limit(DecodeLimit::InputBytes {
            observed: source.len(),
            maximum: options.max_input_bytes,
        }));
    }
    if options.max_nesting == 0 || options.max_nesting > MAX_RECURSION {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: options.max_nesting,
            maximum: MAX_RECURSION,
        }));
    }
    Ok(())
}

fn body_options(source: &[u8], options: DecodeOptions) -> pages_body_codec::DecodeOptions {
    pages_body_codec::DecodeOptions::new(
        source.len().max(1),
        options.max_fields.max(1),
        options
            .max_work_bytes
            .max(source.len().saturating_mul(4))
            .max(1),
        options.max_nesting,
    )
}

fn footnote_options(source: &[u8], options: DecodeOptions) -> pages_footnote_codec::DecodeOptions {
    pages_footnote_codec::DecodeOptions::new(
        source.len().max(1),
        options.max_fields.max(1),
        options
            .max_work_bytes
            .max(source.len().saturating_mul(8))
            .max(1),
        options.max_nesting,
    )
}

fn marker_options(
    source: &[u8],
    options: DecodeOptions,
) -> pages_footnote_marker_codec::DecodeOptions {
    pages_footnote_marker_codec::DecodeOptions::new(
        source.len().max(1),
        options.max_fields.max(1),
        options
            .max_work_bytes
            .max(source.len().saturating_mul(8))
            .max(1),
        options.max_nesting,
    )
}

fn scan_table(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    let (_fields, entries) = scan_table_fields(source, options, budget)?;
    Ok(entries.len())
}

fn scan_table_fields<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(Vec<RawField<'source>>, Vec<ParsedEntry>), DecodeError> {
    let mut fields = Vec::new();
    fields
        .try_reserve(source.len().min(options.max_fields))
        .map_err(|_| DecodeError::allocation(source.len().min(options.max_fields)))?;
    let mut entries = Vec::new();
    entries
        .try_reserve(options.max_entries.min(source.len().max(1)))
        .map_err(|_| DecodeError::allocation(options.max_entries.min(source.len().max(1))))?;
    let mut remaining = source;
    while let Some(field) = parse_field(&mut remaining, 1, budget)? {
        if field.number == TABLE_ENTRIES_FIELD {
            if field.wire != 2 {
                return Err(DecodeError::wire(
                    "body footnote table entry is not length-delimited",
                ));
            }
            let payload = field.payload.ok_or_else(DecodeError::candidate)?;
            let parsed = parse_entry(payload, options, budget)?;
            if entries.len() >= options.max_entries {
                return Err(DecodeError::limit(DecodeLimit::Entries {
                    observed: entries.len().saturating_add(1),
                    maximum: options.max_entries,
                }));
            }
            if entries
                .last()
                .is_some_and(|last: &ParsedEntry| last.character_index >= parsed.character_index)
            {
                return Err(DecodeError::ordering());
            }
            if entries.iter().any(|entry: &ParsedEntry| {
                entry.reference_identifier == parsed.reference_identifier
            }) {
                return Err(DecodeError::duplicate_key(
                    "body footnote reference identifier",
                ));
            }
            entries.push(parsed);
        }
        fields.push(field);
    }
    Ok((fields, entries))
}

fn parse_entry(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ParsedEntry, DecodeError> {
    budget.charge_depth(2)?;
    let mut remaining = source;
    let mut character_index = None;
    let mut reference_identifier = None;
    while let Some(field) = parse_field(&mut remaining, 2, budget)? {
        match field.number {
            BOUNDARY_CHARACTER_INDEX_FIELD => {
                if character_index.is_some() {
                    return Err(DecodeError::duplicate(
                        "TSWP.ObjectAttributeTable.ObjectAttribute.character_index",
                    ));
                }
                if field.wire != 0 || !field.canonical_value {
                    return Err(DecodeError::invalid(
                        "body footnote character index is not canonical",
                    ));
                }
                let value = field.value.ok_or_else(DecodeError::candidate)?;
                character_index = Some(u32::try_from(value).map_err(|_| {
                    DecodeError::invalid("body footnote character index exceeds u32")
                })?);
            },
            BOUNDARY_OBJECT_FIELD => {
                if reference_identifier.is_some() {
                    return Err(DecodeError::duplicate(
                        "TSWP.ObjectAttributeTable.ObjectAttribute.object",
                    ));
                }
                if field.wire != 2 {
                    return Err(DecodeError::wire(
                        "body footnote object reference is not length-delimited",
                    ));
                }
                let payload = field.payload.ok_or_else(DecodeError::candidate)?;
                reference_identifier = Some(parse_reference(payload, 3, budget)?);
            },
            _ => {},
        }
    }
    let expected_index = character_index.ok_or_else(|| {
        DecodeError::missing("TSWP.ObjectAttributeTable.ObjectAttribute.character_index")
    })?;
    let expected_reference = reference_identifier
        .ok_or_else(|| DecodeError::missing("TSWP.ObjectAttributeTable.ObjectAttribute.object"))?;
    // Run the sibling strict Buffa projection as a parity oracle over a
    // canonical known-field projection. Opaque source fields intentionally do
    // not cross that projection: Buffa may reject an overlong unknown scalar
    // even though this source-preserving layer must retain it byte-for-byte.
    let canonical = canonical_entry_payload(expected_index, expected_reference);
    let snapshot =
        pages_body_codec::decode_section_boundary(&canonical, body_options(&canonical, options))
            .map_err(|_| DecodeError::candidate())?;
    if snapshot.character_index() != expected_index
        || snapshot
            .section()
            .is_none_or(|reference| reference.identifier() != expected_reference)
    {
        return Err(DecodeError::candidate());
    }
    Ok(ParsedEntry {
        character_index: expected_index,
        reference_identifier: expected_reference,
    })
}

fn parse_reference(
    source: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<NonZeroU64, DecodeError> {
    budget.charge_depth(depth)?;
    let mut remaining = source;
    let mut identifier = None;
    while let Some(field) = parse_field(&mut remaining, depth, budget)? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {},
            REFERENCE_DEPRECATED_TYPE_FIELD | REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                return Err(DecodeError::invalid(
                    "TSP.Reference deprecated fields are not canonical here",
                ));
            },
            _ => continue,
        }
        if identifier.is_some() {
            return Err(DecodeError::duplicate("TSP.Reference.identifier"));
        }
        if field.wire != 0 || !field.canonical_value {
            return Err(DecodeError::invalid(
                "TSP.Reference.identifier is not canonical",
            ));
        }
        identifier = Some(
            NonZeroU64::new(field.value.ok_or_else(DecodeError::candidate)?)
                .ok_or_else(|| DecodeError::invalid("TSP.Reference.identifier is zero"))?,
        );
    }
    identifier.ok_or_else(|| DecodeError::missing("TSP.Reference.identifier"))
}

fn parse_field<'source>(
    source: &mut &'source [u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<Option<RawField<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    let original = *source;
    let (tag, canonical_key) = take_varint(source)?;
    if !canonical_key {
        return Err(DecodeError::invalid("protobuf field key is not canonical"));
    }
    budget.charge_field()?;
    let raw_tag = u32::try_from(tag).map_err(|_| DecodeError::wire("field number overflow"))?;
    let number = raw_tag >> 3;
    if number == 0 || number > 0x1fff_ffff {
        return Err(DecodeError::wire("invalid protobuf field number"));
    }
    let wire = u8::try_from(raw_tag & 7).map_err(|_| DecodeError::wire("wire type overflow"))?;
    let (payload, value, canonical_value) = match wire {
        0 => {
            let (value, canonical) = take_varint(source)?;
            (None, Some(value), canonical)
        },
        1 => {
            take_exact(source, 8)?;
            (None, None, true)
        },
        2 => {
            let (length, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::invalid(
                    "length-delimited size is not canonical",
                ));
            }
            let length = usize::try_from(length)
                .map_err(|_| DecodeError::wire("length-delimited size overflow"))?;
            (Some(take_exact(source, length)?), None, true)
        },
        3 => {
            let child_depth = depth.checked_add(1).ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Nesting {
                    observed: u32::MAX,
                    maximum: budget.max_nesting,
                })
            })?;
            budget.charge_depth(child_depth)?;
            skip_group(source, number, child_depth, budget)?;
            (None, None, true)
        },
        4 => return Err(DecodeError::wire("unexpected protobuf end-group")),
        5 => {
            take_exact(source, 4)?;
            (None, None, true)
        },
        _ => return Err(DecodeError::wire("invalid protobuf wire type")),
    };
    let consumed = original.len().saturating_sub(source.len());
    let raw = &original[..consumed];
    budget.charge_work(consumed)?;
    Ok(Some(RawField {
        number,
        wire,
        raw,
        payload,
        value,
        canonical_value,
    }))
}

fn skip_group(
    source: &mut &[u8],
    expected: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    loop {
        if source.is_empty() {
            return Err(DecodeError::wire("unterminated protobuf group"));
        }
        let original = *source;
        let (tag, canonical_key) = take_varint(source)?;
        if !canonical_key {
            return Err(DecodeError::invalid("protobuf group key is not canonical"));
        }
        budget.charge_field()?;
        let raw_tag = u32::try_from(tag).map_err(|_| DecodeError::wire("group field overflow"))?;
        let number = raw_tag >> 3;
        let wire =
            u8::try_from(raw_tag & 7).map_err(|_| DecodeError::wire("group wire overflow"))?;
        match wire {
            0 => {
                let _ = take_varint(source)?;
            },
            1 => {
                take_exact(source, 8)?;
            },
            2 => {
                let (length, canonical) = take_varint(source)?;
                if !canonical {
                    return Err(DecodeError::invalid("group length is not canonical"));
                }
                let length = usize::try_from(length)
                    .map_err(|_| DecodeError::wire("group length overflow"))?;
                take_exact(source, length)?;
            },
            3 => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    DecodeError::limit(DecodeLimit::Nesting {
                        observed: u32::MAX,
                        maximum: budget.max_nesting,
                    })
                })?;
                budget.charge_depth(child_depth)?;
                skip_group(source, number, child_depth, budget)?;
            },
            4 if number == expected => {
                let consumed = original.len().saturating_sub(source.len());
                budget.charge_work(consumed)?;
                return Ok(());
            },
            4 => return Err(DecodeError::wire("mismatched protobuf end-group")),
            5 => {
                take_exact(source, 4)?;
            },
            _ => return Err(DecodeError::wire("invalid group wire type")),
        }
        let consumed = original.len().saturating_sub(source.len());
        budget.charge_work(consumed)?;
    }
}

fn take_varint(source: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original
            .get(index)
            .ok_or_else(|| DecodeError::wire("truncated protobuf varint"))?;
        if index == 9 && byte > 1 {
            return Err(DecodeError::wire("protobuf varint is too long"));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            *source = &original[consumed..];
            return Ok((value, canonical_varint_len(value) == consumed));
        }
    }
    Err(DecodeError::wire("protobuf varint is too long"))
}

fn take_exact<'source>(
    source: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], DecodeError> {
    if source.len() < length {
        return Err(DecodeError::wire("truncated protobuf field"));
    }
    let (selected, remaining) = source.split_at(length);
    *source = remaining;
    Ok(selected)
}

fn canonical_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn parse_field_unmetered<'source>(
    source: &mut &'source [u8],
) -> Result<Option<UnmeteredField<'source>>, ()> {
    if source.is_empty() {
        return Ok(None);
    }
    let original = *source;
    let (tag, _) = read_varint_unmetered(source)?;
    let number = u32::try_from(tag >> 3).map_err(|_| ())?;
    if number == 0 || number > 0x1fff_ffff {
        return Err(());
    }
    let payload;
    let value = match (tag & 7) as u8 {
        0 => UnmeteredValue::Varint(read_varint_unmetered(source)?.0),
        1 => {
            take_exact_unmetered(source, 8)?;
            UnmeteredValue::Fixed64
        },
        2 => {
            let length = usize::try_from(read_varint_unmetered(source)?.0).map_err(|_| ())?;
            payload = take_exact_unmetered(source, length)?;
            UnmeteredValue::LengthDelimited(payload)
        },
        3 => {
            skip_group_unmetered(source, number)?;
            UnmeteredValue::Group
        },
        4 => return Err(()),
        5 => {
            take_exact_unmetered(source, 4)?;
            UnmeteredValue::Fixed32
        },
        _ => return Err(()),
    };
    let consumed = original.len().saturating_sub(source.len());
    let raw = &original[..consumed];
    let payload = match value {
        UnmeteredValue::LengthDelimited(payload) => Some(payload),
        _ => None,
    };
    Ok(Some(UnmeteredField {
        number,
        raw,
        payload,
        value,
    }))
}

fn parse_entry_unmetered(source: &[u8]) -> Result<ParsedEntry, ()> {
    let mut remaining = source;
    let mut character_index = None;
    let mut reference_identifier = None;
    while let Some(field) = parse_field_unmetered(&mut remaining)? {
        match (field.number, field.value) {
            (BOUNDARY_CHARACTER_INDEX_FIELD, UnmeteredValue::Varint(value)) => {
                if character_index
                    .replace(u32::try_from(value).map_err(|_| ())?)
                    .is_some()
                {
                    return Err(());
                }
            },
            (BOUNDARY_OBJECT_FIELD, UnmeteredValue::LengthDelimited(payload)) => {
                if reference_identifier
                    .replace(parse_reference_unmetered(payload)?)
                    .is_some()
                {
                    return Err(());
                }
            },
            _ => {},
        }
    }
    Ok(ParsedEntry {
        character_index: character_index.ok_or(())?,
        reference_identifier: reference_identifier.ok_or(())?,
    })
}

fn parse_reference_unmetered(source: &[u8]) -> Result<NonZeroU64, ()> {
    let mut remaining = source;
    let mut identifier = None;
    while let Some(field) = parse_field_unmetered(&mut remaining)? {
        if field.number != REFERENCE_IDENTIFIER_FIELD {
            continue;
        }
        let UnmeteredValue::Varint(value) = field.value else {
            return Err(());
        };
        if identifier
            .replace(NonZeroU64::new(value).ok_or(())?)
            .is_some()
        {
            return Err(());
        }
    }
    identifier.ok_or(())
}

fn read_varint_unmetered(source: &mut &[u8]) -> Result<(u64, bool), ()> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original.get(index).ok_or(())?;
        if index == 9 && byte > 1 {
            return Err(());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            *source = &original[consumed..];
            return Ok((value, canonical_varint_len(value) == consumed));
        }
    }
    Err(())
}

fn take_exact_unmetered<'source>(
    source: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], ()> {
    if source.len() < length {
        return Err(());
    }
    let (selected, remaining) = source.split_at(length);
    *source = remaining;
    Ok(selected)
}

fn skip_group_unmetered(source: &mut &[u8], expected: u32) -> Result<(), ()> {
    loop {
        let (tag, _) = read_varint_unmetered(source)?;
        let number = u32::try_from(tag >> 3).map_err(|_| ())?;
        match (tag & 7) as u8 {
            0 => {
                let _ = read_varint_unmetered(source)?;
            },
            1 => {
                take_exact_unmetered(source, 8)?;
            },
            2 => {
                let length = usize::try_from(read_varint_unmetered(source)?.0).map_err(|_| ())?;
                take_exact_unmetered(source, length)?;
            },
            3 => skip_group_unmetered(source, number)?,
            4 if number == expected => return Ok(()),
            4 => return Err(()),
            5 => {
                take_exact_unmetered(source, 4)?;
            },
            _ => return Err(()),
        }
    }
}

fn validate_requested_entries(
    entries: &[BodyFootnoteEntryWrite<'_>],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    if entries.len() > options.max_entries {
        return Err(DecodeError::limit(DecodeLimit::Entries {
            observed: entries.len(),
            maximum: options.max_entries,
        }));
    }
    let mut previous_index = None;
    for (index, entry) in entries.iter().copied().enumerate() {
        if previous_index.is_some_and(|previous| previous >= entry.character_index) {
            return Err(DecodeError::ordering());
        }
        if entries[..index]
            .iter()
            .any(|previous| previous.reference_identifier == entry.reference_identifier)
        {
            return Err(DecodeError::duplicate_key(
                "body footnote reference identifier",
            ));
        }
        previous_index = Some(entry.character_index);
        budget.charge_work(16)?;
        if entry.preserves_raw() {
            let raw = entry.raw.ok_or_else(DecodeError::candidate)?;
            let parsed = parse_entry_unmetered(raw).map_err(|_| DecodeError::candidate())?;
            if parsed.character_index != entry.character_index
                || parsed.reference_identifier != entry.reference_identifier
            {
                return Err(DecodeError::candidate());
            }
            let field_raw = entry.field_raw.ok_or_else(DecodeError::candidate)?;
            let parsed_field = parse_field_unmetered(&mut &field_raw[..])
                .map_err(|_| DecodeError::candidate())?
                .ok_or_else(DecodeError::candidate)?;
            if parsed_field.number != TABLE_ENTRIES_FIELD || parsed_field.payload != Some(raw) {
                return Err(DecodeError::candidate());
            }
        }
    }
    Ok(())
}

fn measured_table_output(
    fields: &[RawField<'_>],
    source_entries: usize,
    requested: &[BodyFootnoteEntryWrite<'_>],
) -> Result<usize, DecodeError> {
    let unknown_bytes = fields
        .iter()
        .filter(|field| field.number != TABLE_ENTRIES_FIELD)
        .try_fold(0usize, |total, field| total.checked_add(field.raw.len()))
        .ok_or_else(|| DecodeError::wire("body table output size overflow"))?;
    let requested_bytes =
        requested
            .iter()
            .try_fold(0usize, |total, entry| -> Result<usize, DecodeError> {
                let length = entry_field_length(*entry)?;
                total
                    .checked_add(length)
                    .ok_or_else(|| DecodeError::wire("body table output size overflow"))
            })?;
    let _ = source_entries;
    unknown_bytes
        .checked_add(requested_bytes)
        .ok_or_else(|| DecodeError::wire("body table output size overflow"))
}

fn entry_field_length(entry: BodyFootnoteEntryWrite<'_>) -> Result<usize, DecodeError> {
    if let Some(raw) = entry.field_raw {
        return Ok(raw.len());
    }
    let payload = canonical_entry_payload(entry.character_index, entry.reference_identifier);
    field_bytes_len(TABLE_ENTRIES_FIELD, payload.len())
}

fn emit_table(
    fields: &[RawField<'_>],
    source_entries: usize,
    requested: &[BodyFootnoteEntryWrite<'_>],
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut requested_index = 0usize;
    let mut source_entry_index = 0usize;
    for field in fields {
        if field.number != TABLE_ENTRIES_FIELD {
            output.extend_from_slice(field.raw);
            continue;
        }
        if let Some(entry) = requested.get(requested_index).copied() {
            emit_entry(entry, output)?;
            requested_index = requested_index.saturating_add(1);
        }
        source_entry_index = source_entry_index.saturating_add(1);
    }
    if source_entry_index != source_entries {
        return Err(DecodeError::candidate());
    }
    while let Some(entry) = requested.get(requested_index).copied() {
        emit_entry(entry, output)?;
        requested_index = requested_index.saturating_add(1);
    }
    Ok(())
}

fn emit_entry(entry: BodyFootnoteEntryWrite<'_>, output: &mut Vec<u8>) -> Result<(), DecodeError> {
    if let Some(raw) = entry.field_raw {
        output.extend_from_slice(raw);
        return Ok(());
    }
    let payload = canonical_entry_payload(entry.character_index, entry.reference_identifier);
    append_field_bytes(output, TABLE_ENTRIES_FIELD, &payload);
    Ok(())
}

fn verify_candidate_entries(
    candidate: BodyFootnoteTableSnapshot<'_>,
    requested: &[BodyFootnoteEntryWrite<'_>],
) -> Result<(), DecodeError> {
    if candidate.len() != requested.len() {
        return Err(DecodeError::candidate());
    }
    for (actual, expected) in candidate.entries().zip(requested.iter().copied()) {
        if actual.character_index() != expected.character_index
            || actual.reference_identifier() != expected.reference_identifier
        {
            return Err(DecodeError::candidate());
        }
    }
    Ok(())
}

fn reserve_output(amount: usize) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(amount)
        .map_err(|_| DecodeError::allocation(amount))?;
    if output.capacity() != amount {
        return Err(DecodeError::allocation(amount));
    }
    Ok(output)
}

fn validate_graph_write(
    write: FootnoteGraphWrite<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let identifiers = [
        Some(write.reference_identifier()),
        Some(write.storage_identifier()),
        Some(write.marker_identifier()),
        write.stylesheet_identifier(),
        write.paragraph_style_identifier(),
        write.list_style_identifier(),
    ];
    if identifiers.iter().enumerate().any(|(index, identifier)| {
        identifier.is_some() && identifiers[..index].contains(identifier)
    }) {
        return Err(DecodeError::invalid(
            "graph and storage-template object identifiers must be distinct",
        ));
    }
    if write.text().contains(['\u{000e}', '\u{fffc}']) {
        return Err(DecodeError::invalid(
            "footnote text contains a structural marker",
        ));
    }
    let text_bytes = FOOTNOTE_TEXT_PREFIX
        .len()
        .checked_add(write.text().len())
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::TextBytes {
                observed: usize::MAX,
                maximum: options.max_input_bytes,
            })
        })?;
    if text_bytes > options.max_input_bytes {
        return Err(DecodeError::limit(DecodeLimit::TextBytes {
            observed: text_bytes,
            maximum: options.max_input_bytes,
        }));
    }
    if let Some(custom_mark) = write.custom_mark() {
        if custom_mark.contains(['\u{000e}', '\u{fffc}']) {
            return Err(DecodeError::invalid(
                "custom footnote mark contains a structural marker",
            ));
        }
        if custom_mark.len() > options.max_input_bytes {
            return Err(DecodeError::limit(DecodeLimit::TextBytes {
                observed: custom_mark.len(),
                maximum: options.max_input_bytes,
            }));
        }
    }
    Ok(())
}

fn format_text(text: &str) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(FOOTNOTE_TEXT_PREFIX.len().saturating_add(text.len()))
        .map_err(|_| {
            DecodeError::allocation(FOOTNOTE_TEXT_PREFIX.len().saturating_add(text.len()))
        })?;
    output.extend_from_slice(FOOTNOTE_TEXT_PREFIX.as_bytes());
    output.extend_from_slice(text.as_bytes());
    Ok(output)
}

fn canonical_marker_payload() -> Vec<u8> {
    let mut output = Vec::new();
    append_field_varint(&mut output, TEXTUAL_KIND_FIELD, FOOTNOTE_MARK_KIND as u64);
    output
}

fn canonical_reference_payload(
    storage_identifier: NonZeroU64,
    custom_mark: Option<&str>,
) -> Vec<u8> {
    let textual = canonical_textual_payload();
    let reference = canonical_reference(storage_identifier);
    let mut output = Vec::new();
    append_field_bytes(&mut output, FOOTNOTE_REFERENCE_SUPER_FIELD, &textual);
    append_field_bytes(&mut output, FOOTNOTE_REFERENCE_STORAGE_FIELD, &reference);
    if let Some(custom_mark) = custom_mark {
        append_field_bytes(
            &mut output,
            FOOTNOTE_REFERENCE_CUSTOM_MARK_FIELD,
            custom_mark.as_bytes(),
        );
    }
    output
}

fn canonical_textual_payload() -> Vec<u8> {
    let mut output = Vec::new();
    append_field_varint(&mut output, TEXTUAL_KIND_FIELD, FOOTNOTE_MARK_KIND as u64);
    output
}

fn canonical_reference(identifier: NonZeroU64) -> Vec<u8> {
    let mut output = Vec::new();
    append_field_varint(&mut output, REFERENCE_IDENTIFIER_FIELD, identifier.get());
    output
}

fn canonical_reference_size(identifier: NonZeroU64) -> usize {
    canonical_varint_field_size(REFERENCE_IDENTIFIER_FIELD, identifier.get())
}

fn canonical_reference_payload_size(
    storage_identifier: NonZeroU64,
    custom_mark: Option<&str>,
) -> Result<usize, DecodeError> {
    let mut size = 0usize;
    add_size(
        &mut size,
        canonical_bytes_field_size(FOOTNOTE_REFERENCE_SUPER_FIELD, canonical_textual_size())?,
    )?;
    add_size(
        &mut size,
        canonical_bytes_field_size(
            FOOTNOTE_REFERENCE_STORAGE_FIELD,
            canonical_reference_size(storage_identifier),
        )?,
    )?;
    if let Some(custom_mark) = custom_mark {
        add_size(
            &mut size,
            canonical_bytes_field_size(FOOTNOTE_REFERENCE_CUSTOM_MARK_FIELD, custom_mark.len())?,
        )?;
    }
    Ok(size)
}

fn canonical_textual_size() -> usize {
    canonical_varint_field_size(TEXTUAL_KIND_FIELD, FOOTNOTE_MARK_KIND as u64)
}

fn canonical_marker_payload_size() -> usize {
    canonical_textual_size()
}

fn canonical_object_attribute_table_size(
    object_identifier: Option<NonZeroU64>,
) -> Result<usize, DecodeError> {
    let mut entry_size = canonical_varint_field_size(BOUNDARY_CHARACTER_INDEX_FIELD, 0);
    if let Some(identifier) = object_identifier {
        add_size(
            &mut entry_size,
            canonical_bytes_field_size(
                BOUNDARY_OBJECT_FIELD,
                canonical_reference_size(identifier),
            )?,
        )?;
    }
    canonical_bytes_field_size(TABLE_ENTRIES_FIELD, entry_size)
}

fn canonical_para_data_table_size() -> Result<usize, DecodeError> {
    let entry_size = canonical_varint_field_size(BOUNDARY_CHARACTER_INDEX_FIELD, 0)
        .checked_add(canonical_varint_field_size(2, 0))
        .and_then(|value| value.checked_add(canonical_varint_field_size(3, 0)))
        .ok_or_else(|| DecodeError::wire("para-data size overflow"))?;
    canonical_bytes_field_size(TABLE_ENTRIES_FIELD, entry_size)
}

fn canonical_language_table_size(language: &str) -> Result<usize, DecodeError> {
    let entry_size = canonical_varint_field_size(BOUNDARY_CHARACTER_INDEX_FIELD, 0)
        .checked_add(canonical_bytes_field_size(2, language.len())?)
        .ok_or_else(|| DecodeError::wire("language table size overflow"))?;
    canonical_bytes_field_size(TABLE_ENTRIES_FIELD, entry_size)
}

fn canonical_storage_payload_size(
    text_len: usize,
    write: FootnoteGraphWrite<'_>,
) -> Result<usize, DecodeError> {
    let mut size = canonical_varint_field_size(STORAGE_KIND_FIELD, FOOTNOTE_KIND);
    if let Some(identifier) = write.stylesheet_identifier() {
        add_size(
            &mut size,
            canonical_bytes_field_size(
                STORAGE_STYLESHEET_FIELD,
                canonical_reference_size(identifier),
            )?,
        )?;
    }
    add_size(
        &mut size,
        canonical_bytes_field_size(STORAGE_TEXT_FIELD, text_len)?,
    )?;
    add_size(
        &mut size,
        canonical_bytes_field_size(
            STORAGE_PARA_STYLE_FIELD,
            canonical_object_attribute_table_size(write.paragraph_style_identifier())?,
        )?,
    )?;
    let para_data_size = canonical_para_data_table_size()?;
    add_size(
        &mut size,
        canonical_bytes_field_size(STORAGE_PARA_DATA_FIELD, para_data_size)?,
    )?;
    add_size(
        &mut size,
        canonical_bytes_field_size(
            STORAGE_LIST_STYLE_FIELD,
            canonical_object_attribute_table_size(write.list_style_identifier())?,
        )?,
    )?;
    add_size(
        &mut size,
        canonical_bytes_field_size(
            STORAGE_ATTACHMENT_TABLE_FIELD,
            canonical_object_attribute_table_size(Some(write.marker_identifier()))?,
        )?,
    )?;
    add_size(
        &mut size,
        canonical_varint_field_size(STORAGE_IN_DOCUMENT_FIELD, 1),
    )?;
    add_size(
        &mut size,
        canonical_bytes_field_size(STORAGE_PARA_STARTS_FIELD, para_data_size)?,
    )?;
    if let Some(language) = write.language() {
        add_size(
            &mut size,
            canonical_bytes_field_size(
                STORAGE_LANGUAGE_FIELD,
                canonical_language_table_size(language)?,
            )?,
        )?;
    }
    add_size(
        &mut size,
        canonical_bytes_field_size(STORAGE_PARA_BIDI_FIELD, para_data_size)?,
    )?;
    add_size(
        &mut size,
        canonical_bytes_field_size(
            STORAGE_DROP_CAP_FIELD,
            canonical_object_attribute_table_size(None)?,
        )?,
    )?;
    Ok(size)
}

fn canonical_entry_payload_size(
    character_index: u32,
    reference_identifier: NonZeroU64,
) -> Result<usize, DecodeError> {
    canonical_varint_field_size(BOUNDARY_CHARACTER_INDEX_FIELD, u64::from(character_index))
        .checked_add(canonical_bytes_field_size(
            BOUNDARY_OBJECT_FIELD,
            canonical_reference_size(reference_identifier),
        )?)
        .ok_or_else(|| DecodeError::wire("body footnote entry size overflow"))
}

fn canonical_body_entry_payload_size(
    character_index: u32,
    reference_identifier: NonZeroU64,
) -> Result<usize, DecodeError> {
    canonical_bytes_field_size(
        TABLE_ENTRIES_FIELD,
        canonical_entry_payload_size(character_index, reference_identifier)?,
    )
}

fn canonical_varint_field_size(number: u32, value: u64) -> usize {
    varint_len(u64::from(number) << 3) + varint_len(value)
}

fn canonical_bytes_field_size(number: u32, payload_len: usize) -> Result<usize, DecodeError> {
    let payload_len_u64 = u64::try_from(payload_len)
        .map_err(|_| DecodeError::wire("length-delimited field size overflow"))?;
    varint_len((u64::from(number) << 3) | 2)
        .checked_add(varint_len(payload_len_u64))
        .and_then(|value| value.checked_add(payload_len))
        .ok_or_else(|| DecodeError::wire("length-delimited field size overflow"))
}

fn add_size(total: &mut usize, amount: usize) -> Result<(), DecodeError> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| DecodeError::wire("canonical payload size overflow"))?;
    Ok(())
}

fn canonical_storage_payload(text: &[u8], write: FootnoteGraphWrite<'_>) -> Vec<u8> {
    let attachment_table = canonical_storage_marker_table(write.marker_identifier());

    let mut output = Vec::new();
    append_field_varint(&mut output, STORAGE_KIND_FIELD, FOOTNOTE_KIND);
    if let Some(stylesheet_identifier) = write.stylesheet_identifier() {
        let stylesheet = canonical_reference(stylesheet_identifier);
        append_field_bytes(&mut output, STORAGE_STYLESHEET_FIELD, &stylesheet);
    }
    append_field_bytes(&mut output, STORAGE_TEXT_FIELD, text);

    let paragraph_style = canonical_object_attribute_table(write.paragraph_style_identifier());
    append_field_bytes(&mut output, STORAGE_PARA_STYLE_FIELD, &paragraph_style);
    let para_data = canonical_para_data_table();
    append_field_bytes(&mut output, STORAGE_PARA_DATA_FIELD, &para_data);
    let list_style = canonical_object_attribute_table(write.list_style_identifier());
    append_field_bytes(&mut output, STORAGE_LIST_STYLE_FIELD, &list_style);
    append_field_bytes(
        &mut output,
        STORAGE_ATTACHMENT_TABLE_FIELD,
        &attachment_table,
    );
    append_field_varint(&mut output, STORAGE_IN_DOCUMENT_FIELD, 1);
    append_field_bytes(&mut output, STORAGE_PARA_STARTS_FIELD, &para_data);
    if let Some(language) = write.language() {
        let language_table = canonical_language_table(language);
        append_field_bytes(&mut output, STORAGE_LANGUAGE_FIELD, &language_table);
    }
    append_field_bytes(&mut output, STORAGE_PARA_BIDI_FIELD, &para_data);
    let drop_cap = canonical_object_attribute_table(None);
    append_field_bytes(&mut output, STORAGE_DROP_CAP_FIELD, &drop_cap);
    output
}

fn canonical_object_attribute_table(object_identifier: Option<NonZeroU64>) -> Vec<u8> {
    let mut entry = Vec::new();
    append_field_varint(&mut entry, BOUNDARY_CHARACTER_INDEX_FIELD, 0);
    if let Some(identifier) = object_identifier {
        let reference = canonical_reference(identifier);
        append_field_bytes(&mut entry, BOUNDARY_OBJECT_FIELD, &reference);
    }
    let mut output = Vec::new();
    append_field_bytes(&mut output, TABLE_ENTRIES_FIELD, &entry);
    output
}

fn canonical_para_data_table() -> Vec<u8> {
    let mut entry = Vec::new();
    append_field_varint(&mut entry, BOUNDARY_CHARACTER_INDEX_FIELD, 0);
    append_field_varint(&mut entry, 2, 0);
    append_field_varint(&mut entry, 3, 0);
    let mut output = Vec::new();
    append_field_bytes(&mut output, TABLE_ENTRIES_FIELD, &entry);
    output
}

fn canonical_language_table(language: &str) -> Vec<u8> {
    let mut entry = Vec::new();
    append_field_varint(&mut entry, BOUNDARY_CHARACTER_INDEX_FIELD, 0);
    append_field_bytes(&mut entry, 2, language.as_bytes());
    let mut output = Vec::new();
    append_field_bytes(&mut output, TABLE_ENTRIES_FIELD, &entry);
    output
}

fn canonical_storage_marker_table(marker_identifier: NonZeroU64) -> Vec<u8> {
    let marker_reference = canonical_reference(marker_identifier);
    let mut marker_entry = Vec::new();
    append_field_varint(&mut marker_entry, BOUNDARY_CHARACTER_INDEX_FIELD, 0);
    append_field_bytes(&mut marker_entry, BOUNDARY_OBJECT_FIELD, &marker_reference);
    let mut attachment_table = Vec::new();
    append_field_bytes(&mut attachment_table, TABLE_ENTRIES_FIELD, &marker_entry);
    attachment_table
}

fn canonical_graph_field_count(write: FootnoteGraphWrite<'_>) -> usize {
    let reference_fields = 4usize + usize::from(write.custom_mark().is_some());
    let object_table_fields = |object_identifier: Option<NonZeroU64>| {
        2usize + usize::from(object_identifier.is_some()) * 2
    };
    let storage_fields = 1usize
        + usize::from(write.stylesheet_identifier().is_some()) * 2
        + 1
        + object_table_fields(write.paragraph_style_identifier())
        + 4
        + object_table_fields(write.list_style_identifier())
        + 4
        + 1
        + 4
        + usize::from(write.language().is_some()) * 3
        + 4
        + 2;
    let marker_fields = 1usize;
    let body_entry_fields = 4usize;
    reference_fields + storage_fields + marker_fields + body_entry_fields
}

fn verify_storage_payload(
    source: &[u8],
    expected_text: &[u8],
    write: FootnoteGraphWrite<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let mut budget = Budget::new(source, options);
    let mut remaining = source;
    let mut saw_kind = false;
    let mut saw_stylesheet = false;
    let mut saw_text = false;
    let mut saw_para_style = false;
    let mut saw_para_data = false;
    let mut saw_list_style = false;
    let mut saw_attachment = false;
    let mut saw_in_document = false;
    let mut saw_para_starts = false;
    let mut saw_language = false;
    let mut saw_para_bidi = false;
    let mut saw_drop_cap = false;
    while let Some(field) = parse_field(&mut remaining, 1, &mut budget)? {
        match field.number {
            STORAGE_KIND_FIELD => {
                reject_duplicate(&mut saw_kind, "TSWP.StorageArchive.kind")?;
                require_canonical_varint(&field, FOOTNOTE_KIND, "TSWP.StorageArchive.kind")?;
            },
            STORAGE_STYLESHEET_FIELD => {
                reject_duplicate(&mut saw_stylesheet, "TSWP.StorageArchive.style_sheet")?;
                let payload = require_length_delimited(&field, "TSWP.StorageArchive.style_sheet")?;
                let identifier = parse_reference(payload, 2, &mut budget)?;
                if write.stylesheet_identifier() != Some(identifier) {
                    return Err(DecodeError::candidate());
                }
            },
            STORAGE_TEXT_FIELD => {
                reject_duplicate(&mut saw_text, "TSWP.StorageArchive.text")?;
                let payload = require_length_delimited(&field, "TSWP.StorageArchive.text")?;
                if payload != expected_text {
                    return Err(DecodeError::candidate());
                }
            },
            STORAGE_PARA_STYLE_FIELD => {
                reject_duplicate(&mut saw_para_style, "TSWP.StorageArchive.table_para_style")?;
                parse_object_attribute_table(
                    require_length_delimited(&field, "TSWP.StorageArchive.table_para_style")?,
                    write.paragraph_style_identifier(),
                    2,
                    &mut budget,
                )?;
            },
            STORAGE_PARA_DATA_FIELD => {
                reject_duplicate(&mut saw_para_data, "TSWP.StorageArchive.table_para_data")?;
                parse_para_data_table(
                    require_length_delimited(&field, "TSWP.StorageArchive.table_para_data")?,
                    2,
                    &mut budget,
                )?;
            },
            STORAGE_LIST_STYLE_FIELD => {
                reject_duplicate(&mut saw_list_style, "TSWP.StorageArchive.table_list_style")?;
                parse_object_attribute_table(
                    require_length_delimited(&field, "TSWP.StorageArchive.table_list_style")?,
                    write.list_style_identifier(),
                    2,
                    &mut budget,
                )?;
            },
            STORAGE_ATTACHMENT_TABLE_FIELD => {
                reject_duplicate(&mut saw_attachment, "TSWP.StorageArchive.table_attachment")?;
                parse_object_attribute_table(
                    require_length_delimited(&field, "TSWP.StorageArchive.table_attachment")?,
                    Some(write.marker_identifier()),
                    2,
                    &mut budget,
                )?;
            },
            STORAGE_IN_DOCUMENT_FIELD => {
                reject_duplicate(&mut saw_in_document, "TSWP.StorageArchive.in_document")?;
                require_canonical_varint(&field, 1, "TSWP.StorageArchive.in_document")?;
            },
            STORAGE_PARA_STARTS_FIELD => {
                reject_duplicate(
                    &mut saw_para_starts,
                    "TSWP.StorageArchive.table_para_starts",
                )?;
                parse_para_data_table(
                    require_length_delimited(&field, "TSWP.StorageArchive.table_para_starts")?,
                    2,
                    &mut budget,
                )?;
            },
            STORAGE_LANGUAGE_FIELD => {
                reject_duplicate(&mut saw_language, "TSWP.StorageArchive.table_language")?;
                let expected = write.language().ok_or_else(DecodeError::candidate)?;
                parse_language_table(
                    require_length_delimited(&field, "TSWP.StorageArchive.table_language")?,
                    expected.as_bytes(),
                    2,
                    &mut budget,
                )?;
            },
            STORAGE_PARA_BIDI_FIELD => {
                reject_duplicate(&mut saw_para_bidi, "TSWP.StorageArchive.table_para_bidi")?;
                parse_para_data_table(
                    require_length_delimited(&field, "TSWP.StorageArchive.table_para_bidi")?,
                    2,
                    &mut budget,
                )?;
            },
            STORAGE_DROP_CAP_FIELD => {
                reject_duplicate(
                    &mut saw_drop_cap,
                    "TSWP.StorageArchive.table_drop_cap_style",
                )?;
                parse_object_attribute_table(
                    require_length_delimited(&field, "TSWP.StorageArchive.table_drop_cap_style")?,
                    None,
                    2,
                    &mut budget,
                )?;
            },
            _ => {
                return Err(DecodeError::invalid(
                    "canonical storage contains an unknown field",
                ));
            },
        }
    }
    if !saw_kind
        || !saw_text
        || !saw_para_style
        || !saw_para_data
        || !saw_list_style
        || !saw_attachment
        || !saw_in_document
        || !saw_para_starts
        || !saw_para_bidi
        || !saw_drop_cap
    {
        return Err(DecodeError::candidate());
    }
    if saw_stylesheet != write.stylesheet_identifier().is_some()
        || saw_language != write.language().is_some()
    {
        return Err(DecodeError::candidate());
    }
    Ok(())
}

fn reject_duplicate(seen: &mut bool, field: &'static str) -> Result<(), DecodeError> {
    if *seen {
        return Err(DecodeError::duplicate(field));
    }
    *seen = true;
    Ok(())
}

fn require_canonical_varint(
    field: &RawField<'_>,
    expected: u64,
    label: &'static str,
) -> Result<(), DecodeError> {
    if field.wire != 0 || !field.canonical_value || field.value != Some(expected) {
        return Err(DecodeError::invalid(label));
    }
    Ok(())
}

fn require_length_delimited<'source>(
    field: &RawField<'source>,
    label: &'static str,
) -> Result<&'source [u8], DecodeError> {
    if field.wire != 2 {
        return Err(DecodeError::invalid(label));
    }
    field.payload.ok_or_else(DecodeError::candidate)
}

fn parse_object_attribute_table(
    source: &[u8],
    expected_object: Option<NonZeroU64>,
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_depth(depth)?;
    let mut remaining = source;
    let mut entry_seen = false;
    while let Some(field) = parse_field(&mut remaining, depth, budget)? {
        if field.number != TABLE_ENTRIES_FIELD {
            return Err(DecodeError::invalid(
                "canonical object attribute table contains an unknown field",
            ));
        }
        reject_duplicate(&mut entry_seen, "TSWP.ObjectAttributeTable.entries")?;
        let payload = require_length_delimited(&field, "TSWP.ObjectAttributeTable.entries")?;
        parse_object_attribute(payload, expected_object, depth.saturating_add(1), budget)?;
    }
    if !entry_seen {
        return Err(DecodeError::missing("TSWP.ObjectAttributeTable.entries"));
    }
    Ok(())
}

fn parse_object_attribute(
    source: &[u8],
    expected_object: Option<NonZeroU64>,
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_depth(depth)?;
    let mut remaining = source;
    let mut index_seen = false;
    let mut object_seen = false;
    let mut actual_object = None;
    while let Some(field) = parse_field(&mut remaining, depth, budget)? {
        match field.number {
            BOUNDARY_CHARACTER_INDEX_FIELD => {
                reject_duplicate(
                    &mut index_seen,
                    "TSWP.ObjectAttributeTable.ObjectAttribute.character_index",
                )?;
                require_canonical_varint(
                    &field,
                    0,
                    "TSWP.ObjectAttributeTable.ObjectAttribute.character_index",
                )?;
            },
            BOUNDARY_OBJECT_FIELD => {
                reject_duplicate(
                    &mut object_seen,
                    "TSWP.ObjectAttributeTable.ObjectAttribute.object",
                )?;
                let payload = require_length_delimited(
                    &field,
                    "TSWP.ObjectAttributeTable.ObjectAttribute.object",
                )?;
                actual_object = Some(parse_reference(payload, depth.saturating_add(1), budget)?);
            },
            _ => {
                return Err(DecodeError::invalid(
                    "canonical object attribute contains an unknown field",
                ));
            },
        }
    }
    if !index_seen || actual_object != expected_object || object_seen != expected_object.is_some() {
        return Err(DecodeError::candidate());
    }
    Ok(())
}

fn parse_para_data_table(
    source: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_depth(depth)?;
    let mut remaining = source;
    let mut table_entry_seen = false;
    while let Some(field) = parse_field(&mut remaining, depth, budget)? {
        if field.number != TABLE_ENTRIES_FIELD {
            return Err(DecodeError::invalid(
                "canonical para-data table contains an unknown field",
            ));
        }
        reject_duplicate(&mut table_entry_seen, "TSWP.ParaDataAttributeTable.entries")?;
        parse_para_data_entry(
            require_length_delimited(&field, "TSWP.ParaDataAttributeTable.entries")?,
            depth.saturating_add(1),
            budget,
        )?;
    }
    if !table_entry_seen {
        return Err(DecodeError::missing("TSWP.ParaDataAttributeTable.entries"));
    }
    Ok(())
}

fn parse_para_data_entry(
    source: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_depth(depth)?;
    let mut remaining = source;
    let mut seen = [false; 3];
    while let Some(field) = parse_field(&mut remaining, depth, budget)? {
        let slot = match field.number {
            1 => 0,
            2 => 1,
            3 => 2,
            _ => {
                return Err(DecodeError::invalid(
                    "canonical para-data entry contains an unknown field",
                ));
            },
        };
        if seen[slot] {
            return Err(DecodeError::duplicate(
                "TSWP.ParaDataAttributeTable.ParaDataAttribute",
            ));
        }
        seen[slot] = true;
        require_canonical_varint(&field, 0, "TSWP.ParaDataAttributeTable.ParaDataAttribute")?;
    }
    if seen != [true, true, true] {
        return Err(DecodeError::candidate());
    }
    Ok(())
}

fn parse_language_table(
    source: &[u8],
    expected: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_depth(depth)?;
    let mut remaining = source;
    let mut table_entry_seen = false;
    while let Some(field) = parse_field(&mut remaining, depth, budget)? {
        if field.number != TABLE_ENTRIES_FIELD {
            return Err(DecodeError::invalid(
                "canonical language table contains an unknown field",
            ));
        }
        reject_duplicate(&mut table_entry_seen, "TSWP.StringAttributeTable.entries")?;
        parse_language_entry(
            require_length_delimited(&field, "TSWP.StringAttributeTable.entries")?,
            expected,
            depth.saturating_add(1),
            budget,
        )?;
    }
    if !table_entry_seen {
        return Err(DecodeError::missing("TSWP.StringAttributeTable.entries"));
    }
    Ok(())
}

fn parse_language_entry(
    source: &[u8],
    expected: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_depth(depth)?;
    let mut remaining = source;
    let mut index_seen = false;
    let mut object_seen = false;
    while let Some(field) = parse_field(&mut remaining, depth, budget)? {
        match field.number {
            1 => {
                reject_duplicate(
                    &mut index_seen,
                    "TSWP.StringAttributeTable.StringAttribute.character_index",
                )?;
                require_canonical_varint(
                    &field,
                    0,
                    "TSWP.StringAttributeTable.StringAttribute.character_index",
                )?;
            },
            2 => {
                reject_duplicate(
                    &mut object_seen,
                    "TSWP.StringAttributeTable.StringAttribute.object",
                )?;
                if require_length_delimited(
                    &field,
                    "TSWP.StringAttributeTable.StringAttribute.object",
                )? != expected
                {
                    return Err(DecodeError::candidate());
                }
            },
            _ => {
                return Err(DecodeError::invalid(
                    "canonical language entry contains an unknown field",
                ));
            },
        }
    }
    if !index_seen || !object_seen {
        return Err(DecodeError::candidate());
    }
    Ok(())
}

fn canonical_entry_payload(character_index: u32, reference_identifier: NonZeroU64) -> Vec<u8> {
    let reference = canonical_reference(reference_identifier);
    let mut payload = Vec::new();
    append_field_varint(
        &mut payload,
        BOUNDARY_CHARACTER_INDEX_FIELD,
        u64::from(character_index),
    );
    append_field_bytes(&mut payload, BOUNDARY_OBJECT_FIELD, &reference);
    payload
}

fn canonical_body_entry_payload(character_index: u32, reference_identifier: NonZeroU64) -> Vec<u8> {
    let payload = canonical_entry_payload(character_index, reference_identifier);
    let mut output = Vec::new();
    append_field_bytes(&mut output, TABLE_ENTRIES_FIELD, &payload);
    output
}

fn append_field_varint(output: &mut Vec<u8>, number: u32, value: u64) {
    put_varint(output, u64::from(number) << 3);
    put_varint(output, value);
}

fn field_bytes_len(number: u32, payload_len: usize) -> Result<usize, DecodeError> {
    let payload_len = u64::try_from(payload_len)
        .map_err(|_| DecodeError::wire("length-delimited field size overflow"))?;
    varint_len((u64::from(number) << 3) | 2)
        .checked_add(varint_len(payload_len))
        .and_then(|value| value.checked_add(payload_len as usize))
        .ok_or_else(|| DecodeError::wire("length-delimited field size overflow"))
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn append_field_bytes(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    put_varint(output, (u64::from(number) << 3) | 2);
    put_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn put_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}

fn verify_graph_payloads(
    reference: &[u8],
    storage: &[u8],
    marker: &[u8],
    body_entry: &[u8],
    write: FootnoteGraphWrite<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let reference_view = pages_footnote_codec::decode_footnote_reference(
        reference,
        footnote_options(reference, options),
    )
    .map_err(|_| DecodeError::candidate())?;
    if reference_view.super_kind() != Some(FOOTNOTE_MARK_KIND)
        || reference_view
            .contained_storage()
            .is_none_or(|value| value.identifier() != write.storage_identifier())
        || reference_view.custom_mark_string() != write.custom_mark()
    {
        return Err(DecodeError::candidate());
    }
    let marker_view = pages_footnote_marker_codec::decode_textual_attachment(
        marker,
        marker_options(marker, options),
    )
    .map_err(|_| DecodeError::candidate())?;
    if marker_view.kind() != Some(FOOTNOTE_MARK_KIND) {
        return Err(DecodeError::candidate());
    }
    let expected = format_text(write.text())?;
    verify_storage_payload(storage, &expected, write, options)?;
    let mut body_wrapper = body_entry;
    let body_field = parse_field_unmetered(&mut body_wrapper)
        .map_err(|_| DecodeError::candidate())?
        .ok_or_else(DecodeError::candidate)?;
    if !body_wrapper.is_empty()
        || body_field.number != TABLE_ENTRIES_FIELD
        || body_field.payload.is_none()
    {
        return Err(DecodeError::candidate());
    }
    let body_payload = body_field.payload.ok_or_else(DecodeError::candidate)?;
    let body_view = pages_body_codec::decode_section_boundary(
        body_payload,
        body_options(body_payload, options),
    )
    .map_err(|_| DecodeError::candidate())?;
    if body_view.character_index() != write.character_index()
        || body_view
            .section()
            .is_none_or(|value| value.identifier() != write.reference_identifier())
    {
        return Err(DecodeError::candidate());
    }
    let text_view = text_storage_codec::decode_storage_text(
        storage,
        text_storage_codec::DecodeOptions::new(
            storage.len().max(1),
            options.max_fields,
            options.max_work_bytes.max(storage.len().saturating_mul(4)),
            options.max_nesting,
        ),
    )
    .map_err(|_| DecodeError::candidate())?;
    let Some(fragment) = text_view.fragments().next() else {
        return Err(DecodeError::candidate());
    };
    if fragment.as_bytes() != expected.as_slice() || text_view.fragments().count() != 1 {
        return Err(DecodeError::candidate());
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    reason = "Focused graph-codec tests use compact generated-wire fixtures."
)]
mod tests {
    use super::*;
    use crate::{tsp, tswp};
    use prost::Message as _;

    fn identifier(value: u64) -> NonZeroU64 {
        NonZeroU64::new(value).expect("nonzero test identifier")
    }

    fn options() -> DecodeOptions {
        DecodeOptions::new(16 * 1024, 16 * 1024, 512, 1_000_000, 8, 64)
    }

    fn reference(value: NonZeroU64) -> tsp::Reference {
        tsp::Reference {
            identifier: value.get(),
            ..Default::default()
        }
    }

    fn object_attribute_table(object: Option<NonZeroU64>) -> tswp::ObjectAttributeTable {
        tswp::ObjectAttributeTable {
            entries: vec![tswp::object_attribute_table::ObjectAttribute {
                character_index: 0,
                object: object.map(reference),
            }],
        }
    }

    fn para_data_table() -> tswp::ParaDataAttributeTable {
        tswp::ParaDataAttributeTable {
            entries: vec![tswp::para_data_attribute_table::ParaDataAttribute {
                character_index: 0,
                first: 0,
                second: 0,
            }],
        }
    }

    fn language_table(language: &str) -> tswp::StringAttributeTable {
        tswp::StringAttributeTable {
            entries: vec![tswp::string_attribute_table::StringAttribute {
                character_index: 0,
                object: Some(language.to_owned()),
            }],
        }
    }

    fn styled_write<'text>(text: &'text str) -> FootnoteGraphWrite<'text> {
        FootnoteGraphWrite::new(identifier(100), identifier(101), identifier(102), 12, text)
            .with_custom_mark(Some("custom"))
            .with_storage_template(
                Some(identifier(103)),
                Some(identifier(104)),
                Some(identifier(105)),
                Some("en-US"),
            )
    }

    #[test]
    fn canonical_styled_graph_matches_native_storage_template() {
        let write = styled_write("hello — 北区");
        let (payloads, report) =
            encode_footnote_graph_with_report(write, options()).expect("canonical footnote graph");
        let storage = tswp::StorageArchive::decode(payloads.storage()).expect("storage archive");
        let expected_storage = tswp::StorageArchive {
            kind: Some(FOOTNOTE_KIND as i32),
            style_sheet: Some(reference(identifier(103))),
            text: vec![format!("{FOOTNOTE_TEXT_PREFIX}hello — 北区")],
            in_document: Some(true),
            table_para_style: Some(object_attribute_table(Some(identifier(104)))),
            table_para_data: Some(para_data_table()),
            table_list_style: Some(object_attribute_table(Some(identifier(105)))),
            table_attachment: Some(object_attribute_table(Some(identifier(102)))),
            table_para_starts: Some(para_data_table()),
            table_language: Some(language_table("en-US")),
            table_para_bidi: Some(para_data_table()),
            table_drop_cap_style: Some(object_attribute_table(None)),
            ..Default::default()
        };
        assert_eq!(storage, expected_storage);
        assert_eq!(
            payloads.storage(),
            expected_storage.encode_to_vec().as_slice()
        );

        let reference_archive =
            tswp::FootnoteReferenceAttachmentArchive::decode(payloads.reference())
                .expect("reference archive");
        assert_eq!(
            reference_archive
                .super_
                .as_ref()
                .and_then(|value| value.kind),
            Some(FOOTNOTE_MARK_KIND),
        );
        assert_eq!(
            reference_archive
                .contained_storage
                .as_ref()
                .map(|value| value.identifier),
            Some(identifier(101).get()),
        );
        assert_eq!(
            reference_archive.custom_mark_string.as_deref(),
            Some("custom")
        );

        let marker_archive =
            tswp::TextualAttachmentArchive::decode(payloads.marker()).expect("marker archive");
        assert_eq!(marker_archive.kind, Some(FOOTNOTE_MARK_KIND));

        let body_table =
            tswp::ObjectAttributeTable::decode(payloads.body_entry()).expect("body entry table");
        assert_eq!(body_table.entries.len(), 1);
        assert_eq!(body_table.entries[0].character_index, 12);
        assert_eq!(
            body_table.entries[0]
                .object
                .as_ref()
                .map(|value| value.identifier),
            Some(identifier(100).get()),
        );
        assert_eq!(
            report.output_bytes(),
            report.reference_bytes()
                + report.storage_bytes()
                + report.marker_bytes()
                + report.body_entry_bytes()
        );
        assert_eq!(report.fields(), canonical_graph_field_count(write));
    }

    #[test]
    fn graph_report_accepts_exact_budget_and_rejects_each_minus_one() {
        let write = styled_write("bounded");
        let baseline = encode_footnote_graph_with_report(write, options()).expect("baseline");
        let report = baseline.1;
        let exact = DecodeOptions::new(
            16 * 1024,
            report.output_bytes(),
            report.fields(),
            report.work_bytes(),
            8,
            64,
        );
        encode_footnote_graph(write, exact).expect("exact graph budget");

        assert!(matches!(
            encode_footnote_graph(
                write,
                exact.with_max_output_bytes(report.output_bytes().saturating_sub(1)),
            ),
            Err(DecodeError {
                kind: DecodeErrorKind::Limit(DecodeLimit::OutputBytes { .. }),
            })
        ));
        assert!(matches!(
            encode_footnote_graph(
                write,
                DecodeOptions::new(
                    16 * 1024,
                    report.output_bytes(),
                    report.fields().saturating_sub(1),
                    report.work_bytes(),
                    8,
                    64,
                ),
            ),
            Err(DecodeError {
                kind: DecodeErrorKind::Limit(DecodeLimit::Fields { .. }),
            })
        ));
        assert!(matches!(
            encode_footnote_graph(
                write,
                DecodeOptions::new(
                    16 * 1024,
                    report.output_bytes(),
                    report.fields(),
                    report.work_bytes().saturating_sub(1),
                    8,
                    64,
                ),
            ),
            Err(DecodeError {
                kind: DecodeErrorKind::Limit(DecodeLimit::WorkBytes { .. }),
            })
        ));
    }

    fn unknown_overlong_scalar(output: &mut Vec<u8>, number: u32) {
        put_varint(output, u64::from(number) << 3);
        output.extend_from_slice(&[0x80, 0x00]);
    }

    fn unknown_group(output: &mut Vec<u8>, number: u32) {
        put_varint(output, (u64::from(number) << 3) | 3);
        unknown_overlong_scalar(output, number + 1);
        put_varint(output, (u64::from(number) << 3) | 4);
    }

    fn raw_entry(character_index: u32, reference_identifier: NonZeroU64) -> Vec<u8> {
        canonical_entry_payload(character_index, reference_identifier)
    }

    fn append_entry_table(output: &mut Vec<u8>, payload: &[u8]) {
        append_field_bytes(output, TABLE_ENTRIES_FIELD, payload);
    }

    #[test]
    fn table_rewrite_preserves_unknown_scalars_groups_and_order() {
        let first = raw_entry(2, identifier(201));
        let second = raw_entry(9, identifier(202));
        let mut source = Vec::new();
        unknown_overlong_scalar(&mut source, 40);
        unknown_group(&mut source, 41);
        append_entry_table(&mut source, &first);
        unknown_group(&mut source, 42);
        append_entry_table(&mut source, &second);

        let (snapshot, decode_report) =
            decode_body_footnote_table_with_report(&source, options()).expect("source table");
        assert_eq!(snapshot.len(), 2);
        assert_eq!(decode_report.entries(), 2);
        let retained_first = BodyFootnoteEntryWrite::preserve(
            snapshot.entries().next().expect("first source entry"),
        );
        let retained_second = BodyFootnoteEntryWrite::preserve(
            snapshot.entries().nth(1).expect("second source entry"),
        );
        let inserted = BodyFootnoteEntryWrite::new(5, identifier(203));
        let requested = [retained_first, inserted, retained_second];
        let (rewritten, report) = rewrite_body_footnote_table_with_report(
            &source,
            BodyFootnoteTableWrite::new(&requested),
            options(),
        )
        .expect("inserted table");
        assert_eq!(report.entries_before(), 2);
        assert_eq!(report.entries_after(), 3);
        assert_eq!(report.inserted(), 1);
        assert_eq!(report.removed(), 0);
        assert!(rewritten.windows(2).any(|window| window == [0x80, 0x00]));
        assert!(
            rewritten
                .windows(4)
                .any(|window| window == [0xd8, 0x02, 0x80, 0x00])
        );
        assert!(
            rewritten
                .windows(first.len())
                .any(|window| window == first.as_slice())
        );
        assert!(
            rewritten
                .windows(second.len())
                .any(|window| window == second.as_slice())
        );

        let retained_only = [retained_first];
        let (removed, removed_report) = rewrite_body_footnote_table_with_report(
            &source,
            BodyFootnoteTableWrite::new(&retained_only),
            options(),
        )
        .expect("removed table");
        assert_eq!(removed_report.entries_before(), 2);
        assert_eq!(removed_report.entries_after(), 1);
        assert_eq!(removed_report.removed(), 1);
        assert!(
            removed
                .windows(first.len())
                .any(|window| window == first.as_slice())
        );
        assert!(
            !removed
                .windows(second.len())
                .any(|window| window == second.as_slice())
        );
        assert!(removed.windows(2).any(|window| window == [0x80, 0x00]));
    }

    #[test]
    fn malformed_known_fields_and_references_are_rejected() {
        let mut noncanonical_index = Vec::new();
        noncanonical_index.extend_from_slice(&[0x08, 0x80, 0x00]);
        let reference = canonical_reference(identifier(301));
        append_field_bytes(&mut noncanonical_index, BOUNDARY_OBJECT_FIELD, &reference);
        let mut source = Vec::new();
        append_entry_table(&mut source, &noncanonical_index);
        assert!(decode_body_footnote_table(&source, options()).is_err());

        let mut zero_reference = Vec::new();
        append_field_varint(&mut zero_reference, BOUNDARY_CHARACTER_INDEX_FIELD, 1);
        let mut zero_reference_payload = Vec::new();
        append_field_varint(&mut zero_reference_payload, REFERENCE_IDENTIFIER_FIELD, 0);
        append_field_bytes(
            &mut zero_reference,
            BOUNDARY_OBJECT_FIELD,
            &zero_reference_payload,
        );
        let mut zero_source = Vec::new();
        append_entry_table(&mut zero_source, &zero_reference);
        assert!(decode_body_footnote_table(&zero_source, options()).is_err());

        let mut deprecated_reference = canonical_reference(identifier(302));
        append_field_varint(
            &mut deprecated_reference,
            REFERENCE_DEPRECATED_TYPE_FIELD,
            17,
        );
        let mut deprecated_entry = Vec::new();
        append_field_varint(&mut deprecated_entry, BOUNDARY_CHARACTER_INDEX_FIELD, 1);
        append_field_bytes(
            &mut deprecated_entry,
            BOUNDARY_OBJECT_FIELD,
            &deprecated_reference,
        );
        let mut deprecated_source = Vec::new();
        append_entry_table(&mut deprecated_source, &deprecated_entry);
        assert!(decode_body_footnote_table(&deprecated_source, options()).is_err());

        let mut mismatched_group = Vec::new();
        put_varint(&mut mismatched_group, (60_u64 << 3) | 3);
        put_varint(&mut mismatched_group, (61_u64 << 3) | 4);
        assert!(decode_body_footnote_table(&mismatched_group, options()).is_err());
    }
}
