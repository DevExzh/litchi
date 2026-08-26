//! Strict borrowed Numbers `TSD.CommentStorageArchive` ingress.
//!
//! The handwritten wire pass is the authority for canonical protobuf
//! validation, aggregate resource limits, and source-order reply streaming.
//! A private Buffa lazy projection is forced only for parity on the selected
//! scalar and singular nested values; it contains no generated reply vector.
//! Every borrowed value points into the caller-owned payload, which remains
//! the only preservation representation.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict semantic preflight intentionally precedes its wire reader."
)]

use core::{fmt, mem::size_of, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_comment_storage_generated::LitchiIwaCommentStorageProjection as projection;

const TEXT_FIELD: u32 = 1;
const CREATION_DATE_FIELD: u32 = 2;
const AUTHOR_FIELD: u32 = 3;
const REPLIES_FIELD: u32 = 4;
const STORAGE_UUID_FIELD: u32 = 5;
const DATE_SECONDS_FIELD: u32 = 1;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const UUID_LOWER_FIELD: u32 = 1;
const UUID_UPPER_FIELD: u32 = 2;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Finite aggregate policy for one comment-storage payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_references: usize,
    max_text_bytes: usize,
}

impl DecodeOptions {
    /// Construct an explicit bytes/fields/work/nesting/reference/text policy.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_references: usize,
        max_text_bytes: usize,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_references,
            max_text_bytes,
        }
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Exact aggregate consumption for one successful strict decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    replies: usize,
    reference_bytes: usize,
    text_bytes: usize,
}

impl DecodeReport {
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
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
    pub const fn references(self) -> usize {
        self.references
    }

    /// Repeated `TSD.CommentStorageArchive.replies` occurrences.
    ///
    /// This is deliberately separate from [`Self::references`], which also
    /// includes the optional `author` reference. Keeping the cardinalities
    /// distinct lets host adapters verify that a streaming visitor observed
    /// every reply without mistaking the author for a reply.
    #[must_use]
    pub const fn replies(self) -> usize {
        self.replies
    }

    #[must_use]
    pub const fn reference_bytes(self) -> usize {
        self.reference_bytes
    }

    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    #[must_use]
    pub const fn reply_references(self) -> usize {
        self.replies
    }
}

/// Typed finite resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    Bytes {
        observed: usize,
        maximum: usize,
    },
    /// Candidate output bytes exceeded the prepared execution ceiling.
    OutputBytes {
        observed: usize,
        maximum: usize,
    },
    References {
        observed: usize,
        maximum: usize,
    },
    /// Repeated `replies` occurrences exceeded the prepared execution ceiling.
    Replies {
        observed: usize,
        maximum: usize,
    },
    /// Bytes in reference envelopes exceeded the prepared execution ceiling.
    ReferenceBytes {
        observed: usize,
        maximum: usize,
    },
    Text {
        observed: usize,
        maximum: usize,
    },
    Fields {
        observed: usize,
        maximum: usize,
    },
    Work {
        observed: usize,
        maximum: usize,
    },
    Nesting {
        observed: u32,
        maximum: u32,
    },
    /// The prepared operation required more logical allocations than allowed.
    Allocations {
        observed: usize,
        maximum: usize,
    },
    /// Temporary rewrite scratch exceeded the prepared execution ceiling.
    Scratch {
        observed: usize,
        maximum: usize,
    },
    /// Retained candidate bytes exceeded the prepared execution ceiling.
    Retained {
        observed: usize,
        maximum: usize,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Resource(DecodeLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    InvalidUtf8(&'static str),
    Invalid,
}

/// Strict raw-wire or private Buffa parity failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

impl DecodeError {
    const fn invalid() -> Self {
        Self {
            kind: DecodeErrorKind::Invalid,
        }
    }

    const fn resource(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Resource(limit),
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

    const fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason),
        }
    }

    const fn utf8(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::InvalidUtf8(field),
        }
    }

    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        if let DecodeErrorKind::Resource(limit) = self.kind {
            Some(limit)
        } else {
            None
        }
    }

    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        if let DecodeErrorKind::MissingRequired(field) = self.kind {
            Some(field)
        } else {
            None
        }
    }

    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        if let DecodeErrorKind::DuplicateSingular(field) = self.kind {
            Some(field)
        } else {
            None
        }
    }

    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        if let DecodeErrorKind::NonCanonical(reason) = self.kind {
            Some(reason)
        } else {
            None
        }
    }

    #[must_use]
    pub const fn invalid_utf8_field(&self) -> Option<&'static str> {
        if let DecodeErrorKind::InvalidUtf8(field) = self.kind {
            Some(field)
        } else {
            None
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Resource(DecodeLimit::Bytes { .. }) => {
                formatter.write_str("Numbers comment-storage byte limit exceeded")
            },
            DecodeErrorKind::Resource(DecodeLimit::OutputBytes { .. }) => {
                formatter.write_str("Numbers comment-storage output-byte limit exceeded")
            },
            DecodeErrorKind::Resource(DecodeLimit::References { .. }) => {
                formatter.write_str("Numbers comment-storage reference limit exceeded")
            },
            DecodeErrorKind::Resource(DecodeLimit::Replies { .. }) => {
                formatter.write_str("Numbers comment-storage reply limit exceeded")
            },
            DecodeErrorKind::Resource(DecodeLimit::ReferenceBytes { .. }) => {
                formatter.write_str("Numbers comment-storage reference-byte limit exceeded")
            },
            DecodeErrorKind::Resource(DecodeLimit::Text { .. }) => {
                formatter.write_str("Numbers comment-storage text limit exceeded")
            },
            DecodeErrorKind::Resource(DecodeLimit::Fields { .. }) => {
                formatter.write_str("Numbers comment-storage field limit exceeded")
            },
            DecodeErrorKind::Resource(DecodeLimit::Work { .. }) => {
                formatter.write_str("Numbers comment-storage work limit exceeded")
            },
            DecodeErrorKind::Resource(DecodeLimit::Nesting { .. }) => {
                formatter.write_str("Numbers comment-storage nesting limit exceeded")
            },
            DecodeErrorKind::Resource(DecodeLimit::Allocations { .. }) => {
                formatter.write_str("Numbers comment-storage allocation limit exceeded")
            },
            DecodeErrorKind::Resource(DecodeLimit::Scratch { .. }) => {
                formatter.write_str("Numbers comment-storage scratch limit exceeded")
            },
            DecodeErrorKind::Resource(DecodeLimit::Retained { .. }) => {
                formatter.write_str("Numbers comment-storage retained-byte limit exceeded")
            },
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::InvalidUtf8(field) => {
                write!(formatter, "{field} is invalid UTF-8")
            },
            DecodeErrorKind::Invalid => {
                formatter.write_str("invalid Numbers comment-storage payload")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        match error {
            buffa::DecodeError::MessageTooLarge => Self::resource(DecodeLimit::Bytes {
                observed: 0,
                maximum: 0,
            }),
            buffa::DecodeError::RecursionLimitExceeded => Self::resource(DecodeLimit::Nesting {
                observed: 0,
                maximum: 0,
            }),
            error => Self {
                kind: DecodeErrorKind::Wire(error),
            },
        }
    }
}

/// Presence-preserving IEEE-754 date scalar.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DateSnapshot {
    seconds_bits: u64,
}

impl DateSnapshot {
    #[must_use]
    pub const fn from_bits(seconds_bits: u64) -> Self {
        Self { seconds_bits }
    }

    #[must_use]
    pub const fn seconds_bits(self) -> u64 {
        self.seconds_bits
    }

    #[must_use]
    pub const fn seconds(self) -> f64 {
        f64::from_bits(self.seconds_bits)
    }
}

/// Generated-free scalar projection of one canonical `TSP.Reference`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReferenceSnapshot {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

impl ReferenceSnapshot {
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }

    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }

    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }
}

/// Generated-free scalar projection of one canonical `TSP.UUID`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct UuidSnapshot {
    lower: u64,
    upper: u64,
}

impl UuidSnapshot {
    #[must_use]
    pub const fn from_parts(lower: u64, upper: u64) -> Self {
        Self { lower, upper }
    }

    #[must_use]
    pub const fn lower(self) -> u64 {
        self.lower
    }

    #[must_use]
    pub const fn upper(self) -> u64 {
        self.upper
    }
}

/// One source-ordered repeated reply reference.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct ReferenceRecord<'source> {
    raw: &'source [u8],
    reference: ReferenceSnapshot,
}

impl<'source> ReferenceRecord<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }

    #[must_use]
    pub const fn reference(self) -> ReferenceSnapshot {
        self.reference
    }

    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.reference.identifier()
    }
}

impl fmt::Debug for ReferenceRecord<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ReferenceRecord")
            .field("raw", &"<borrowed>")
            .field("reference", &self.reference)
            .finish()
    }
}

/// Borrowed scalar and singular facts from one comment-storage payload.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct CommentStorageSnapshot<'source> {
    text: Option<&'source str>,
    creation_date: Option<DateSnapshot>,
    author: Option<ReferenceSnapshot>,
    storage_uuid: Option<UuidSnapshot>,
}

impl<'source> CommentStorageSnapshot<'source> {
    #[must_use]
    pub const fn text(self) -> Option<&'source str> {
        self.text
    }

    #[must_use]
    pub const fn creation_date(self) -> Option<DateSnapshot> {
        self.creation_date
    }

    #[must_use]
    pub const fn author(self) -> Option<ReferenceSnapshot> {
        self.author
    }

    #[must_use]
    pub const fn storage_uuid(self) -> Option<UuidSnapshot> {
        self.storage_uuid
    }
}

impl fmt::Debug for CommentStorageSnapshot<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("CommentStorageSnapshot")
            .field("text", &self.text)
            .field("creation_date", &self.creation_date)
            .field("author", &self.author)
            .field("storage_uuid", &self.storage_uuid)
            .finish()
    }
}

/// Streaming hook for source-ordered `TSD.CommentStorageArchive.replies`.
///
/// A callback can observe validated records before the enclosing root has
/// completed parity and validation. A later wire, parity, or callback error
/// does not roll back earlier calls; callers should stage side effects until
/// the decode returns `Ok`.
pub trait CommentStorageVisitor {
    fn visit_reply(&mut self, _reply: ReferenceRecord<'_>) -> Result<(), DecodeError> {
        Ok(())
    }
}

impl CommentStorageVisitor for () {}

/// Decode one comment-storage payload without retaining repeated replies.
pub fn decode_comment_storage_archive(
    source: &[u8],
    options: DecodeOptions,
) -> Result<CommentStorageSnapshot<'_>, DecodeError> {
    Ok(decode_comment_storage_archive_with_report(source, options)?.0)
}

/// Decode one canonical `TSP.Reference` without materializing a generated
/// message.  This narrow helper is shared by adjacent comment metadata
/// readers whose source bytes remain authoritative for preservation.
pub fn decode_reference(
    source: &[u8],
    options: DecodeOptions,
) -> Result<ReferenceSnapshot, DecodeError> {
    let mut budget = Budget::new(source, options)?;
    budget.reference(source.len(), false)?;
    let reference = decode_reference_fields(source, &mut budget, 1)?;
    parity_reference(source, reference, &mut budget, 1)?;
    Ok(reference)
}

/// Decode one comment-storage payload and return aggregate resource usage.
pub fn decode_comment_storage_archive_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(CommentStorageSnapshot<'_>, DecodeReport), DecodeError> {
    decode_comment_storage_archive_with_visitor(source, options, &mut ())
}

/// Decode one comment-storage payload and stream every reply in source order.
pub fn decode_comment_storage_archive_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn CommentStorageVisitor,
) -> Result<(CommentStorageSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_root(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

/// Compatibility alias emphasizing the visitor's reply stream.
pub fn decode_comment_storage_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn CommentStorageVisitor,
) -> Result<(CommentStorageSnapshot<'source>, DecodeReport), DecodeError> {
    decode_comment_storage_archive_with_visitor(source, options, visitor)
}

/// Stream replies through a closure without exposing a generated collection.
pub fn visit_comment_storage_replies<'source, F>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut F,
) -> Result<(CommentStorageSnapshot<'source>, DecodeReport), DecodeError>
where
    F: for<'reply> FnMut(ReferenceRecord<'reply>) -> Result<(), DecodeError>,
{
    struct ClosureVisitor<'visitor, F: ?Sized>(&'visitor mut F);
    impl<F> CommentStorageVisitor for ClosureVisitor<'_, F>
    where
        F: FnMut(ReferenceRecord<'_>) -> Result<(), DecodeError> + ?Sized,
    {
        fn visit_reply(&mut self, reply: ReferenceRecord<'_>) -> Result<(), DecodeError> {
            (self.0)(reply)
        }
    }
    let mut closure = ClosureVisitor(visitor);
    decode_comment_storage_archive_with_visitor(source, options, &mut closure)
}

fn decode_root<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
    visitor: &mut dyn CommentStorageVisitor,
) -> Result<CommentStorageSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut text = None;
    let mut creation_date = None;
    let mut raw_creation_date = None;
    let mut author = None;
    let mut raw_author = None;
    let mut storage_uuid = None;
    let mut raw_storage_uuid = None;
    let mut remaining = source;

    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            TEXT_FIELD => {
                require_decoded_known_field(field, 2, "TSD.CommentStorageArchive.text")?;
                if text.is_some() {
                    return Err(DecodeError::duplicate("TSD.CommentStorageArchive.text"));
                }
                let value = field.bytes()?;
                text = Some(strict_utf8(
                    value,
                    budget,
                    "TSD.CommentStorageArchive.text",
                )?);
            },
            CREATION_DATE_FIELD => {
                require_decoded_known_field(field, 2, "TSD.CommentStorageArchive.creation_date")?;
                if creation_date.is_some() {
                    return Err(DecodeError::duplicate(
                        "TSD.CommentStorageArchive.creation_date",
                    ));
                }
                let raw = field.bytes()?;
                let value = decode_date(raw, budget, child_depth)?;
                creation_date = Some(value);
                raw_creation_date = Some(raw);
            },
            AUTHOR_FIELD => {
                require_decoded_known_field(field, 2, "TSD.CommentStorageArchive.author")?;
                if author.is_some() {
                    return Err(DecodeError::duplicate("TSD.CommentStorageArchive.author"));
                }
                let raw = field.bytes()?;
                budget.reference(raw.len(), false)?;
                let value = decode_reference_fields(raw, budget, child_depth)?;
                author = Some(value);
                raw_author = Some(raw);
            },
            REPLIES_FIELD => {
                require_decoded_known_field(field, 2, "TSD.CommentStorageArchive.replies")?;
                let raw = field.bytes()?;
                budget.reference(raw.len(), true)?;
                let reference = decode_reference_fields(raw, budget, child_depth)?;
                parity_reference(raw, reference, budget, child_depth)?;
                visitor.visit_reply(ReferenceRecord { raw, reference })?;
            },
            STORAGE_UUID_FIELD => {
                require_decoded_known_field(field, 2, "TSD.CommentStorageArchive.storage_uuid")?;
                if storage_uuid.is_some() {
                    return Err(DecodeError::duplicate(
                        "TSD.CommentStorageArchive.storage_uuid",
                    ));
                }
                let raw = field.bytes()?;
                let value = decode_uuid(raw, budget, child_depth)?;
                storage_uuid = Some(value);
                raw_storage_uuid = Some(raw);
            },
            _ => {},
        }
    }

    let snapshot = CommentStorageSnapshot {
        text,
        creation_date,
        author,
        storage_uuid,
    };

    // Buffa's root view is lazy. Charge and force every selected child so the
    // parity pass is covered by the same aggregate work/depth policy.
    budget.message(source, depth)?;
    let view: projection::CommentStorageArchiveLazyView<'source> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let projected = force_root_projection(
        &view,
        budget,
        child_depth,
        raw_creation_date,
        raw_author,
        raw_storage_uuid,
    )?;
    if projected != snapshot {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

fn force_root_projection<'source>(
    view: &projection::CommentStorageArchiveLazyView<'source>,
    budget: &mut Budget,
    child_depth: u32,
    raw_creation_date: Option<&'source [u8]>,
    raw_author: Option<&'source [u8]>,
    raw_storage_uuid: Option<&'source [u8]>,
) -> Result<CommentStorageSnapshot<'source>, DecodeError> {
    let creation_date = match (view.creation_date.get()?, raw_creation_date) {
        (Some(date), Some(raw)) => {
            budget.message(raw, child_depth)?;
            Some(force_date_projection(&date)?)
        },
        (None, None) => None,
        _ => return Err(DecodeError::invalid()),
    };
    let author = match (view.author.get()?, raw_author) {
        (Some(author), Some(raw)) => {
            budget.message(raw, child_depth)?;
            Some(force_reference_projection(&author)?)
        },
        (None, None) => None,
        _ => return Err(DecodeError::invalid()),
    };
    let storage_uuid = match (view.storage_uuid.get()?, raw_storage_uuid) {
        (Some(uuid), Some(raw)) => {
            budget.message(raw, child_depth)?;
            Some(force_uuid_projection(&uuid)?)
        },
        (None, None) => None,
        _ => return Err(DecodeError::invalid()),
    };
    Ok(CommentStorageSnapshot {
        text: view.text,
        creation_date,
        author,
        storage_uuid,
    })
}

fn parity_reference(
    source: &[u8],
    strict: ReferenceSnapshot,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let view: projection::ReferenceLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    if force_reference_projection(&view)? != strict {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn force_date_projection(view: &projection::DateLazyView<'_>) -> Result<DateSnapshot, DecodeError> {
    if !view.has_seconds() {
        return Err(DecodeError::missing("TSP.Date.seconds"));
    }
    Ok(DateSnapshot::from_bits(view.seconds.to_bits()))
}

fn force_reference_projection(
    view: &projection::ReferenceLazyView<'_>,
) -> Result<ReferenceSnapshot, DecodeError> {
    if !view.has_identifier() {
        return Err(DecodeError::missing("TSP.Reference.identifier"));
    }
    Ok(ReferenceSnapshot {
        identifier: view.identifier,
        deprecated_type: view.deprecated_type,
        deprecated_is_external: view.deprecated_is_external,
    })
}

fn force_uuid_projection(view: &projection::UuidLazyView<'_>) -> Result<UuidSnapshot, DecodeError> {
    if !view.has_lower() {
        return Err(DecodeError::missing("TSP.UUID.lower"));
    }
    if !view.has_upper() {
        return Err(DecodeError::missing("TSP.UUID.upper"));
    }
    Ok(UuidSnapshot::from_parts(view.lower, view.upper))
}

fn decode_date(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<DateSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let mut seconds = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number == DATE_SECONDS_FIELD {
            require_decoded_known_field(field, 1, "TSP.Date.seconds")?;
            if seconds.is_some() {
                return Err(DecodeError::duplicate("TSP.Date.seconds"));
            }
            seconds = Some(field.fixed64()?);
        }
    }
    seconds
        .map(DateSnapshot::from_bits)
        .ok_or_else(|| DecodeError::missing("TSP.Date.seconds"))
}

fn decode_reference_fields(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<ReferenceSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                require_decoded_known_field(field, 0, "TSP.Reference.identifier")?;
                if identifier.is_some() {
                    return Err(DecodeError::duplicate("TSP.Reference.identifier"));
                }
                identifier = Some(field.varint()?);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                require_decoded_known_field(field, 0, "TSP.Reference.deprecated_type")?;
                if deprecated_type.is_some() {
                    return Err(DecodeError::duplicate("TSP.Reference.deprecated_type"));
                }
                deprecated_type = Some(canonical_int32(field.varint()?)?);
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                require_decoded_known_field(field, 0, "TSP.Reference.deprecated_is_external")?;
                if deprecated_is_external.is_some() {
                    return Err(DecodeError::duplicate(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                deprecated_is_external = Some(canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    Ok(ReferenceSnapshot {
        identifier: identifier.ok_or_else(|| DecodeError::missing("TSP.Reference.identifier"))?,
        deprecated_type,
        deprecated_is_external,
    })
}

fn decode_uuid(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<UuidSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let mut lower = None;
    let mut upper = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            UUID_LOWER_FIELD => {
                require_decoded_known_field(field, 0, "TSP.UUID.lower")?;
                if lower.is_some() {
                    return Err(DecodeError::duplicate("TSP.UUID.lower"));
                }
                lower = Some(field.varint()?);
            },
            UUID_UPPER_FIELD => {
                require_decoded_known_field(field, 0, "TSP.UUID.upper")?;
                if upper.is_some() {
                    return Err(DecodeError::duplicate("TSP.UUID.upper"));
                }
                upper = Some(field.varint()?);
            },
            _ => {},
        }
    }
    Ok(UuidSnapshot::from_parts(
        lower.ok_or_else(|| DecodeError::missing("TSP.UUID.lower"))?,
        upper.ok_or_else(|| DecodeError::missing("TSP.UUID.upper"))?,
    ))
}

fn strict_utf8<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    field: &'static str,
) -> Result<&'source str, DecodeError> {
    let value = str::from_utf8(source).map_err(|_error| DecodeError::utf8(field))?;
    budget.text(source.len())?;
    Ok(value)
}

fn canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}

fn canonical_int32(value: u64) -> Result<i32, DecodeError> {
    if let Ok(value) = i32::try_from(value) {
        return Ok(value);
    }
    if value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical(
            "int32 scalar is not a sign-extended 32-bit value",
        ));
    }
    i32::try_from(i64::from_ne_bytes(value.to_ne_bytes()))
        .map_err(|_error| DecodeError::noncanonical("int32 scalar is out of range"))
}

#[derive(Clone, Copy, Debug)]
struct Field<'source> {
    number: u32,
    wire_type: u8,
    value: Value<'source>,
    key_canonical: bool,
    value_canonical: bool,
    length_canonical: bool,
}

impl<'source> Field<'source> {
    fn varint(self) -> Result<u64, DecodeError> {
        match (self.wire_type, self.value) {
            (0, Value::Varint(value)) => Ok(value),
            _ => Err(DecodeError::invalid()),
        }
    }

    fn fixed64(self) -> Result<u64, DecodeError> {
        match (self.wire_type, self.value) {
            (1, Value::Fixed64(value)) => Ok(value),
            _ => Err(DecodeError::invalid()),
        }
    }

    fn bytes(self) -> Result<&'source [u8], DecodeError> {
        match (self.wire_type, self.value) {
            (2, Value::Bytes(value)) => Ok(value),
            _ => Err(DecodeError::invalid()),
        }
    }
}

fn require_decoded_known_field(
    field: Field<'_>,
    wire_type: u8,
    name: &'static str,
) -> Result<(), DecodeError> {
    if field.wire_type != wire_type {
        return Err(DecodeError::invalid());
    }
    if !field.key_canonical {
        return Err(DecodeError::noncanonical(name));
    }
    match wire_type {
        0 if !field.value_canonical => Err(DecodeError::noncanonical(name)),
        2 if !field.length_canonical => Err(DecodeError::noncanonical(name)),
        _ => Ok(()),
    }
}

#[derive(Clone, Copy, Debug)]
enum Value<'source> {
    Varint(u64),
    Fixed64(u64),
    Bytes(&'source [u8]),
    Group,
    Fixed32,
}

enum ParseItem<'source> {
    Field(Field<'source>),
    EndGroup(u32),
}

fn next_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<Field<'source>>, DecodeError> {
    match parse_field(source, budget, depth)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(number)) => {
            Err(buffa::DecodeError::InvalidEndGroup(number).into())
        },
        None => Ok(None),
    }
}

fn parse_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    budget.observe_depth(depth)?;
    budget.field()?;
    let (tag, key_canonical) = take_varint(source)?;
    let raw_number = tag >> 3;
    let number =
        u32::try_from(raw_number).map_err(|_error| buffa::DecodeError::InvalidFieldNumber)?;
    if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let wire_type = u8::try_from(tag & 7).map_err(|_error| DecodeError::invalid())?;
    let mut value_canonical = true;
    let mut length_canonical = true;
    let value = match wire_type {
        0 => {
            let (value, canonical) = take_varint(source)?;
            value_canonical = canonical;
            Value::Varint(value)
        },
        1 => Value::Fixed64(u64::from_le_bytes(
            take_exact(source, 8)?
                .try_into()
                .map_err(|_error| DecodeError::invalid())?,
        )),
        2 => {
            let (length, canonical) = take_varint(source)?;
            length_canonical = canonical;
            let length =
                usize::try_from(length).map_err(|_error| buffa::DecodeError::MessageTooLarge)?;
            Value::Bytes(take_exact(source, length)?)
        },
        3 => {
            let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
            skip_group(source, number, budget, child_depth)?;
            Value::Group
        },
        4 => return Ok(Some(ParseItem::EndGroup(number))),
        5 => {
            let _ = take_exact(source, 4)?;
            Value::Fixed32
        },
        _ => {
            return Err(buffa::DecodeError::InvalidWireType(u32::from(wire_type)).into());
        },
    };
    Ok(Some(ParseItem::Field(Field {
        number,
        wire_type,
        value,
        key_canonical,
        value_canonical,
        length_canonical,
    })))
}

fn skip_group(
    source: &mut &[u8],
    expected: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    loop {
        match parse_field(source, budget, depth)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected => return Ok(()),
            Some(ParseItem::EndGroup(number)) => {
                return Err(buffa::DecodeError::InvalidEndGroup(number).into());
            },
            None => return Err(buffa::DecodeError::UnexpectedEof.into()),
        }
    }
}

fn take_exact<'source>(
    source: &mut &'source [u8],
    amount: usize,
) -> Result<&'source [u8], DecodeError> {
    if source.len() < amount {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    let (selected, remaining) = source.split_at(amount);
    *source = remaining;
    Ok(selected)
}

fn take_varint(source: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original
            .get(index)
            .ok_or(buffa::DecodeError::UnexpectedEof)?;
        if index == 9 && byte > 1 {
            return Err(buffa::DecodeError::VarintTooLong.into());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index.checked_add(1).ok_or_else(DecodeError::invalid)?;
            *source = &original[consumed..];
            return Ok((value, encoded_varint_len(value) == consumed));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}

const fn encoded_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

struct Budget {
    options: DecodeOptions,
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    replies: usize,
    reference_bytes: usize,
    text_bytes: usize,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        let hard_bytes =
            usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| DecodeError::invalid())?;
        if options.max_message_bytes > hard_bytes {
            return Err(DecodeError::resource(DecodeLimit::Bytes {
                observed: options.max_message_bytes,
                maximum: hard_bytes,
            }));
        }
        if source.len() > options.max_message_bytes {
            return Err(DecodeError::resource(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }));
        }
        if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
            return Err(DecodeError::resource(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION_LIMIT,
            }));
        }
        Ok(Self {
            options,
            source_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            references: 0,
            replies: 0,
            reference_bytes: 0,
            text_bytes: 0,
        })
    }

    fn message(&mut self, source: &[u8], depth: u32) -> Result<(), DecodeError> {
        if source.len() > self.options.max_message_bytes {
            return Err(DecodeError::resource(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: self.options.max_message_bytes,
            }));
        }
        self.observe_depth(depth)?;
        self.work(source.len())
    }

    fn field(&mut self) -> Result<(), DecodeError> {
        let observed = self
            .fields
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_fields {
            return Err(DecodeError::resource(DecodeLimit::Fields {
                observed,
                maximum: self.options.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn work(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self
            .work_bytes
            .checked_add(amount)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_work_bytes {
            return Err(DecodeError::resource(DecodeLimit::Work {
                observed,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn reference(&mut self, bytes: usize, is_reply: bool) -> Result<(), DecodeError> {
        let observed = self
            .references
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_references {
            return Err(DecodeError::resource(DecodeLimit::References {
                observed,
                maximum: self.options.max_references,
            }));
        }
        self.references = observed;
        if is_reply {
            self.replies = self
                .replies
                .checked_add(1)
                .ok_or_else(DecodeError::invalid)?;
        }
        self.reference_bytes = self
            .reference_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::invalid)?;
        Ok(())
    }

    fn text(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self
            .text_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_text_bytes {
            return Err(DecodeError::resource(DecodeLimit::Text {
                observed,
                maximum: self.options.max_text_bytes,
            }));
        }
        self.text_bytes = observed;
        Ok(())
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError::resource(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    const fn report(&self) -> DecodeReport {
        DecodeReport {
            source_bytes: self.source_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            replies: self.replies,
            reference_bytes: self.reference_bytes,
            text_bytes: self.text_bytes,
        }
    }
}

/// One ordered mutation of `TSD.CommentStorageArchive.replies`.
///
/// The ordinal is the source order among reply fields, not a text position or
/// a generated-message index.  This distinction is important for comments
/// containing equal reply text: reply identities are the only selected wire
/// values and are matched independently of any text projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CommentStorageReplyRewrite {
    /// Append one canonical reply reference after the existing root fields.
    Append { identifier: u64 },
    /// Replace one reply at `ordinal` when its identifier is still `expected_identifier`.
    Replace {
        ordinal: usize,
        expected_identifier: u64,
        replacement_identifier: u64,
    },
    /// Remove one reply at `ordinal` when its identifier is still `expected_identifier`.
    Remove {
        ordinal: usize,
        expected_identifier: u64,
    },
}

impl CommentStorageReplyRewrite {
    /// Construct an append operation.
    #[must_use]
    pub const fn append(identifier: u64) -> Self {
        Self::Append { identifier }
    }

    /// Construct a checked ordinal replacement.
    #[must_use]
    pub const fn replace(
        ordinal: usize,
        expected_identifier: u64,
        replacement_identifier: u64,
    ) -> Self {
        Self::Replace {
            ordinal,
            expected_identifier,
            replacement_identifier,
        }
    }

    /// Construct a checked ordinal removal.
    #[must_use]
    pub const fn remove(ordinal: usize, expected_identifier: u64) -> Self {
        Self::Remove {
            ordinal,
            expected_identifier,
        }
    }
}

/// Aggregate requirements for a prepared comment-reply rewrite.
///
/// The counters include the source preflight and the candidate verification
/// pass.  They intentionally include no package or archive bytes: those
/// physical ceilings belong to the enclosing package transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    /// Source bytes inspected during preparation.
    pub input_bytes: usize,
    /// Exact candidate bytes emitted by execution.
    pub output_bytes: usize,
    /// Source plus candidate wire fields inspected.
    pub fields: usize,
    /// Source plus candidate wire work.
    pub work_bytes: usize,
    /// Maximum source/candidate wire depth.
    pub max_depth: u32,
    /// Source plus candidate reference envelopes.
    pub references: usize,
    /// Source plus candidate reply occurrences.
    pub replies: usize,
    /// Source plus candidate reference payload bytes.
    pub reference_bytes: usize,
    /// Conservative count of logical vector allocations requested by staging,
    /// output, and verification. This is not allocator-call telemetry.
    pub allocations: usize,
    /// Conservative upper bound for requested temporary reply-id/group bytes.
    /// Allocator capacity and package-owned scratch are intentionally outside
    /// this codec-local value.
    pub scratch_bytes: usize,
    /// Requested candidate bytes retained by the returned output. This is the
    /// exact logical length, not a claim about allocator capacity.
    pub retained_bytes: usize,
}

impl RewriteExecutionRequirements {
    /// Return a limit set that accepts exactly these requirements.
    #[must_use]
    pub const fn exact(self) -> RewriteExecutionLimits {
        RewriteExecutionLimits {
            input_bytes: self.input_bytes,
            output_bytes: self.output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            replies: self.replies,
            reference_bytes: self.reference_bytes,
            allocations: self.allocations,
            scratch_bytes: self.scratch_bytes,
            retained_bytes: self.retained_bytes,
        }
    }

    /// Compatibility spelling used by package transaction owners.
    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        self.exact()
    }
}

/// Caller-provided ceilings replayed before a candidate allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    pub input_bytes: usize,
    pub output_bytes: usize,
    pub fields: usize,
    pub work_bytes: usize,
    pub max_depth: u32,
    pub references: usize,
    pub replies: usize,
    pub reference_bytes: usize,
    pub allocations: usize,
    pub scratch_bytes: usize,
    pub retained_bytes: usize,
}

impl RewriteExecutionLimits {
    /// Build exact limits from prepared requirements.
    #[must_use]
    pub const fn exact(requirements: RewriteExecutionRequirements) -> Self {
        requirements.exact()
    }

    /// Start with no additional execution restriction.
    #[must_use]
    pub const fn unrestricted() -> Self {
        Self {
            input_bytes: usize::MAX,
            output_bytes: usize::MAX,
            fields: usize::MAX,
            work_bytes: usize::MAX,
            max_depth: u32::MAX,
            references: usize::MAX,
            replies: usize::MAX,
            reference_bytes: usize::MAX,
            allocations: usize::MAX,
            scratch_bytes: usize::MAX,
            retained_bytes: usize::MAX,
        }
    }

    #[must_use]
    pub const fn with_input_bytes(mut self, value: usize) -> Self {
        self.input_bytes = value;
        self
    }
    #[must_use]
    pub const fn with_output_bytes(mut self, value: usize) -> Self {
        self.output_bytes = value;
        self
    }
    #[must_use]
    pub const fn with_fields(mut self, value: usize) -> Self {
        self.fields = value;
        self
    }
    #[must_use]
    pub const fn with_work_bytes(mut self, value: usize) -> Self {
        self.work_bytes = value;
        self
    }
    #[must_use]
    pub const fn with_max_depth(mut self, value: u32) -> Self {
        self.max_depth = value;
        self
    }
    #[must_use]
    pub const fn with_references(mut self, value: usize) -> Self {
        self.references = value;
        self
    }
    #[must_use]
    pub const fn with_replies(mut self, value: usize) -> Self {
        self.replies = value;
        self
    }
    #[must_use]
    pub const fn with_reference_bytes(mut self, value: usize) -> Self {
        self.reference_bytes = value;
        self
    }
    #[must_use]
    pub const fn with_allocations(mut self, value: usize) -> Self {
        self.allocations = value;
        self
    }
    #[must_use]
    pub const fn with_scratch_bytes(mut self, value: usize) -> Self {
        self.scratch_bytes = value;
        self
    }
    #[must_use]
    pub const fn with_retained_bytes(mut self, value: usize) -> Self {
        self.retained_bytes = value;
        self
    }
}

/// Exact source/candidate accounting returned by a prepared rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    source: DecodeReport,
    result: DecodeReport,
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    replies: usize,
    reference_bytes: usize,
    allocations: usize,
    scratch_bytes: usize,
    retained_bytes: usize,
}

impl RewriteReport {
    #[must_use]
    pub const fn source(self) -> DecodeReport {
        self.source
    }
    #[must_use]
    pub const fn result(self) -> DecodeReport {
        self.result
    }
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
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
    pub const fn references(self) -> usize {
        self.references
    }
    #[must_use]
    pub const fn replies(self) -> usize {
        self.replies
    }
    #[must_use]
    pub const fn reference_bytes(self) -> usize {
        self.reference_bytes
    }
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }
}

/// Candidate bytes and exact accounting for one prepared rewrite.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RewriteOutput {
    bytes: Vec<u8>,
    report: RewriteReport,
}

impl RewriteOutput {
    /// Borrow candidate bytes without transferring ownership.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Alias for callers that use the generic codec output convention.
    #[must_use]
    pub fn output(&self) -> &[u8] {
        self.bytes()
    }

    /// Transfer candidate bytes to the caller.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    /// Transfer candidate bytes to the caller.
    #[must_use]
    pub fn into_output(self) -> Vec<u8> {
        self.into_bytes()
    }

    /// Return exact execution accounting.
    #[must_use]
    pub const fn report(&self) -> RewriteReport {
        self.report
    }
}

#[derive(Debug, Clone, Copy)]
struct RawField {
    number: u32,
    wire_type: u8,
    start: usize,
    key_end: usize,
    value_start: usize,
    payload_start: usize,
    payload_end: usize,
    end: usize,
    key_canonical: bool,
    value_canonical: bool,
    length_canonical: bool,
}

impl RawField {
    const fn raw_len(self) -> Result<usize, DecodeError> {
        match self.end.checked_sub(self.start) {
            Some(length) => Ok(length),
            None => Err(DecodeError::invalid()),
        }
    }

    const fn payload_len(self) -> Result<usize, DecodeError> {
        match self.payload_end.checked_sub(self.payload_start) {
            Some(length) => Ok(length),
            None => Err(DecodeError::invalid()),
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct RawReferenceFacts {
    identifier: u64,
    identifier_field: RawField,
    fields: usize,
}

#[derive(Debug)]
struct RawScanSummary {
    report: DecodeReport,
    reply_ids: Vec<u64>,
    target: Option<RawField>,
    target_reference: Option<RawReferenceFacts>,
    max_group_depth: u32,
}

struct RewriteScanBudget {
    options: DecodeOptions,
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    replies: usize,
    reference_bytes: usize,
    text_bytes: usize,
}

impl RewriteScanBudget {
    fn new(source: &[u8], options: DecodeOptions, input_limit: usize) -> Result<Self, DecodeError> {
        let hard_bytes =
            usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| DecodeError::invalid())?;
        if options.max_message_bytes > hard_bytes {
            return Err(DecodeError::resource(DecodeLimit::Bytes {
                observed: options.max_message_bytes,
                maximum: hard_bytes,
            }));
        }
        if source.len() > hard_bytes {
            return Err(DecodeError::resource(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: hard_bytes,
            }));
        }
        if source.len() > input_limit {
            return Err(DecodeError::resource(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: input_limit,
            }));
        }
        if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
            return Err(DecodeError::resource(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION_LIMIT,
            }));
        }
        Ok(Self {
            options,
            source_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            references: 0,
            replies: 0,
            reference_bytes: 0,
            text_bytes: 0,
        })
    }

    fn message(&mut self, source: &[u8], depth: u32) -> Result<(), DecodeError> {
        self.observe_depth(depth)?;
        self.work(source.len())
    }

    fn field(&mut self) -> Result<(), DecodeError> {
        let observed = self
            .fields
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_fields {
            return Err(DecodeError::resource(DecodeLimit::Fields {
                observed,
                maximum: self.options.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn work(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self
            .work_bytes
            .checked_add(amount)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_work_bytes {
            return Err(DecodeError::resource(DecodeLimit::Work {
                observed,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn reference(&mut self, bytes: usize, is_reply: bool) -> Result<(), DecodeError> {
        let observed = self
            .references
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_references {
            return Err(DecodeError::resource(DecodeLimit::References {
                observed,
                maximum: self.options.max_references,
            }));
        }
        self.references = observed;
        if is_reply {
            self.replies = self
                .replies
                .checked_add(1)
                .ok_or_else(DecodeError::invalid)?;
        }
        self.reference_bytes = self
            .reference_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::invalid)?;
        Ok(())
    }

    fn text(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self
            .text_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_text_bytes {
            return Err(DecodeError::resource(DecodeLimit::Text {
                observed,
                maximum: self.options.max_text_bytes,
            }));
        }
        self.text_bytes = observed;
        Ok(())
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError::resource(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    const fn report(&self) -> DecodeReport {
        DecodeReport {
            source_bytes: self.source_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            replies: self.replies,
            reference_bytes: self.reference_bytes,
            text_bytes: self.text_bytes,
        }
    }
}

/// Prepared reply mutation.  The source and all selected offsets are borrowed
/// from the caller; no candidate output is allocated until `execute` passes
/// every requested execution ceiling.
#[derive(Debug)]
pub struct PreparedCommentStorageReplyRewrite<'source> {
    source: &'source [u8],
    rewrite: CommentStorageReplyRewrite,
    options: DecodeOptions,
    source_report: DecodeReport,
    requirements: RewriteExecutionRequirements,
    reply_ids: Vec<u64>,
    target: Option<RawField>,
    target_reference: Option<RawReferenceFacts>,
    source_group_depth: u32,
}

impl<'source> PreparedCommentStorageReplyRewrite<'source> {
    /// Return exact source preflight accounting.
    #[must_use]
    pub const fn prepare_report(&self) -> DecodeReport {
        self.source_report
    }

    /// Return all candidate execution requirements.
    #[must_use]
    pub const fn execution_requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Compatibility alias for callers that name this value `requirements`.
    #[must_use]
    pub const fn requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute after all physical and wire ceilings have been charged.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_rewrite_limits(self.requirements, limits)?;

        let mut output = Vec::new();
        output
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::resource(DecodeLimit::Allocations {
                    observed: self.requirements.allocations,
                    maximum: limits.allocations,
                })
            })?;
        emit_reply_rewrite(
            &mut output,
            self.source,
            self.rewrite,
            self.target,
            self.target_reference,
            self.requirements.output_bytes,
        )?;
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::invalid());
        }

        let candidate_options = DecodeOptions {
            max_message_bytes: output.len().max(1),
            ..self.options
        };
        let candidate = scan_comment_storage_raw(&output, candidate_options, None)?;
        if !reply_ids_match(&candidate.reply_ids, &self.reply_ids, self.rewrite) {
            return Err(DecodeError::invalid());
        }
        let (_, expected_report) = measure_reply_rewrite(
            self.source,
            self.source_report,
            self.rewrite,
            self.target,
            self.target_reference,
        )?;
        let candidate_with_depth_upper_bound = DecodeReport {
            max_depth: expected_report.max_depth,
            ..candidate.report
        };
        if candidate.report.max_depth > expected_report.max_depth
            || candidate_with_depth_upper_bound != expected_report
        {
            return Err(DecodeError::invalid());
        }
        let report = RewriteReport {
            source: self.source_report,
            result: candidate.report,
            input_bytes: self.source.len(),
            output_bytes: output.len(),
            fields: self.requirements.fields,
            work_bytes: self.requirements.work_bytes,
            max_depth: self.requirements.max_depth,
            references: self.requirements.references,
            replies: self.requirements.replies,
            reference_bytes: self.requirements.reference_bytes,
            allocations: self.requirements.allocations,
            scratch_bytes: self.requirements.scratch_bytes,
            retained_bytes: self.requirements.retained_bytes,
        };
        let _ = self.source_group_depth;
        Ok(RewriteOutput {
            bytes: output,
            report,
        })
    }
}

/// Prepare an ordered comment-reply mutation without allocating candidate bytes.
pub fn prepare_comment_storage_reply_rewrite<'source>(
    source: &'source [u8],
    rewrite: CommentStorageReplyRewrite,
    options: DecodeOptions,
) -> Result<PreparedCommentStorageReplyRewrite<'source>, DecodeError> {
    // This check intentionally precedes every Vec or parser helper.  A caller
    // can therefore use max-input as a hard allocation-free ingress gate.
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: options.max_message_bytes,
        }));
    }
    validate_reply_rewrite(rewrite)?;
    let target_ordinal = match rewrite {
        CommentStorageReplyRewrite::Append { .. } => None,
        CommentStorageReplyRewrite::Replace { ordinal, .. }
        | CommentStorageReplyRewrite::Remove { ordinal, .. } => Some(ordinal),
    };
    let source_scan = scan_comment_storage_raw(source, options, target_ordinal)?;
    let (target, target_reference) = match rewrite {
        CommentStorageReplyRewrite::Append { .. } => (None, None),
        CommentStorageReplyRewrite::Replace {
            ordinal,
            expected_identifier,
            ..
        }
        | CommentStorageReplyRewrite::Remove {
            ordinal,
            expected_identifier,
        } => {
            let target = source_scan.target.ok_or_else(DecodeError::invalid)?;
            let target_reference = source_scan
                .target_reference
                .ok_or_else(DecodeError::invalid)?;
            // The scan records only the requested ordinal.  Keep this explicit
            // check adjacent to the operation so stale callers cannot mutate
            // another occurrence after an ordinal shift.
            if target_reference.identifier != expected_identifier
                || ordinal >= source_scan.reply_ids.len()
            {
                return Err(DecodeError::invalid());
            }
            (Some(target), Some(target_reference))
        },
    };
    validate_reply_identifier_transition(&source_scan.reply_ids, rewrite)?;
    let (output_bytes, result_report) = measure_reply_rewrite(
        source,
        source_scan.report,
        rewrite,
        target,
        target_reference,
    )?;
    let hard_bytes =
        usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| DecodeError::invalid())?;
    if output_bytes > hard_bytes {
        return Err(DecodeError::resource(DecodeLimit::OutputBytes {
            observed: output_bytes,
            maximum: hard_bytes,
        }));
    }
    if result_report.max_depth > options.recursion_limit {
        return Err(DecodeError::resource(DecodeLimit::Nesting {
            observed: result_report.max_depth,
            maximum: options.recursion_limit,
        }));
    }
    if result_report.fields > options.max_fields {
        return Err(DecodeError::resource(DecodeLimit::Fields {
            observed: result_report.fields,
            maximum: options.max_fields,
        }));
    }
    if result_report.work_bytes > options.max_work_bytes {
        return Err(DecodeError::resource(DecodeLimit::Work {
            observed: result_report.work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    if result_report.references > options.max_references {
        return Err(DecodeError::resource(DecodeLimit::References {
            observed: result_report.references,
            maximum: options.max_references,
        }));
    }
    let result_ids_len = result_report.replies;
    let source_reply_scratch = source_scan
        .reply_ids
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or_else(DecodeError::invalid)?;
    let candidate_reply_scratch = result_ids_len
        .checked_mul(size_of::<u64>())
        .ok_or_else(DecodeError::invalid)?;
    let group_depth =
        usize::try_from(source_scan.max_group_depth).map_err(|_error| DecodeError::invalid())?;
    let group_scratch = group_depth
        .checked_mul(size_of::<u32>())
        .and_then(|bytes| bytes.checked_mul(2))
        .ok_or_else(DecodeError::invalid)?;
    let scratch_bytes = source_reply_scratch
        .checked_add(candidate_reply_scratch)
        .and_then(|bytes| bytes.checked_add(group_scratch))
        .ok_or_else(DecodeError::invalid)?;
    let planning_allocations = usize::from(!source_scan.reply_ids.is_empty())
        .checked_add(usize::from(source_scan.max_group_depth != 0))
        .ok_or_else(DecodeError::invalid)?;
    let candidate_allocations = usize::from(result_ids_len != 0)
        .checked_add(usize::from(source_scan.max_group_depth != 0))
        .ok_or_else(DecodeError::invalid)?;
    let allocations = planning_allocations
        .checked_add(1)
        .and_then(|count| count.checked_add(candidate_allocations))
        .ok_or_else(DecodeError::invalid)?;
    let max_depth = source_scan.report.max_depth.max(
        if matches!(rewrite, CommentStorageReplyRewrite::Append { .. }) {
            2
        } else {
            0
        },
    );
    let fields = source_scan
        .report
        .fields
        .checked_add(result_report.fields)
        .ok_or_else(DecodeError::invalid)?;
    let work_bytes = source_scan
        .report
        .work_bytes
        .checked_add(result_report.work_bytes)
        .ok_or_else(DecodeError::invalid)?;
    let references = source_scan
        .report
        .references
        .checked_add(result_report.references)
        .ok_or_else(DecodeError::invalid)?;
    let replies = source_scan
        .report
        .replies
        .checked_add(result_report.replies)
        .ok_or_else(DecodeError::invalid)?;
    let reference_bytes = source_scan
        .report
        .reference_bytes
        .checked_add(result_report.reference_bytes)
        .ok_or_else(DecodeError::invalid)?;
    let requirements = RewriteExecutionRequirements {
        input_bytes: source.len(),
        output_bytes,
        fields,
        work_bytes,
        max_depth,
        references,
        replies,
        reference_bytes,
        allocations,
        scratch_bytes,
        retained_bytes: output_bytes,
    };
    Ok(PreparedCommentStorageReplyRewrite {
        source,
        rewrite,
        options,
        source_report: source_scan.report,
        requirements,
        reply_ids: source_scan.reply_ids,
        target,
        target_reference,
        source_group_depth: source_scan.max_group_depth,
    })
}

fn validate_reply_rewrite(rewrite: CommentStorageReplyRewrite) -> Result<(), DecodeError> {
    let invalid_identifier = |identifier: u64| {
        if identifier == 0 {
            Err(DecodeError::invalid())
        } else {
            Ok(())
        }
    };
    match rewrite {
        CommentStorageReplyRewrite::Append { identifier } => invalid_identifier(identifier),
        CommentStorageReplyRewrite::Replace {
            expected_identifier,
            replacement_identifier,
            ..
        } => {
            invalid_identifier(expected_identifier)?;
            invalid_identifier(replacement_identifier)
        },
        CommentStorageReplyRewrite::Remove {
            expected_identifier,
            ..
        } => invalid_identifier(expected_identifier),
    }
}

fn validate_reply_identifier_transition(
    source: &[u64],
    rewrite: CommentStorageReplyRewrite,
) -> Result<(), DecodeError> {
    for (index, identifier) in source.iter().enumerate() {
        if source
            .get(index.checked_add(1).ok_or_else(DecodeError::invalid)?..)
            .is_some_and(|remaining| remaining.contains(identifier))
        {
            return Err(DecodeError::invalid());
        }
    }
    match rewrite {
        CommentStorageReplyRewrite::Append { identifier } => {
            if source.contains(&identifier) {
                return Err(DecodeError::invalid());
            }
        },
        CommentStorageReplyRewrite::Replace {
            ordinal,
            expected_identifier,
            replacement_identifier,
        } => {
            if replacement_identifier != expected_identifier
                && source.iter().enumerate().any(|(index, identifier)| {
                    index != ordinal && *identifier == replacement_identifier
                })
            {
                return Err(DecodeError::invalid());
            }
        },
        CommentStorageReplyRewrite::Remove { .. } => {},
    }
    Ok(())
}

fn raw_varint(
    source: &[u8],
    offset: usize,
    limit: usize,
) -> Result<(u64, usize, bool), DecodeError> {
    let mut value = 0u64;
    let available = source.len().min(limit);
    for index in 0..10usize {
        let position = offset.checked_add(index).ok_or_else(DecodeError::invalid)?;
        let byte = *source.get(position).ok_or_else(DecodeError::invalid)?;
        if position >= available {
            return Err(DecodeError::invalid());
        }
        if index == 9 && byte > 1 {
            return Err(DecodeError::invalid());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let length = index.checked_add(1).ok_or_else(DecodeError::invalid)?;
            return Ok((value, length, encoded_varint_len(value) == length));
        }
    }
    Err(DecodeError::invalid())
}

fn parse_raw_field(source: &[u8], offset: usize, limit: usize) -> Result<RawField, DecodeError> {
    let (tag, key_length, key_canonical) = raw_varint(source, offset, limit)?;
    let key_end = offset
        .checked_add(key_length)
        .ok_or_else(DecodeError::invalid)?;
    let number = u32::try_from(tag >> 3).map_err(|_error| DecodeError::invalid())?;
    if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid());
    }
    let wire_type = u8::try_from(tag & 7).map_err(|_error| DecodeError::invalid())?;
    let mut field = RawField {
        number,
        wire_type,
        start: offset,
        key_end,
        value_start: key_end,
        payload_start: key_end,
        payload_end: key_end,
        end: key_end,
        key_canonical,
        value_canonical: true,
        length_canonical: true,
    };
    match wire_type {
        0 => {
            let (value, value_length, value_canonical) = raw_varint(source, key_end, limit)?;
            let end = key_end
                .checked_add(value_length)
                .ok_or_else(DecodeError::invalid)?;
            field.value_start = key_end;
            field.payload_start = key_end;
            field.payload_end = end;
            field.end = end;
            field.value_canonical = value_canonical;
            let _ = value;
        },
        1 => {
            let end = key_end.checked_add(8).ok_or_else(DecodeError::invalid)?;
            if end > limit || end > source.len() {
                return Err(DecodeError::invalid());
            }
            field.payload_end = end;
            field.end = end;
        },
        2 => {
            let (length, length_width, length_canonical) = raw_varint(source, key_end, limit)?;
            let payload_start = key_end
                .checked_add(length_width)
                .ok_or_else(DecodeError::invalid)?;
            let payload_length =
                usize::try_from(length).map_err(|_error| DecodeError::invalid())?;
            let payload_end = payload_start
                .checked_add(payload_length)
                .ok_or_else(DecodeError::invalid)?;
            if payload_end > limit || payload_end > source.len() {
                return Err(DecodeError::invalid());
            }
            field.value_start = key_end;
            field.payload_start = payload_start;
            field.payload_end = payload_end;
            field.end = payload_end;
            field.length_canonical = length_canonical;
        },
        3 | 4 => {},
        5 => {
            let end = key_end.checked_add(4).ok_or_else(DecodeError::invalid)?;
            if end > limit || end > source.len() {
                return Err(DecodeError::invalid());
            }
            field.payload_end = end;
            field.end = end;
        },
        _ => return Err(DecodeError::invalid()),
    }
    Ok(field)
}

fn consume_raw_group(
    source: &[u8],
    offset: usize,
    limit: usize,
    expected: u32,
    budget: &mut RewriteScanBudget,
    depth: u32,
) -> Result<(usize, u32), DecodeError> {
    let mut stack = Vec::new();
    stack.try_reserve(1).map_err(|_error| {
        DecodeError::resource(DecodeLimit::Allocations {
            observed: 1,
            maximum: 0,
        })
    })?;
    stack.push(expected);
    let mut cursor = offset;
    let mut max_depth = depth;
    while let Some(&group) = stack.last() {
        if cursor >= limit {
            return Err(DecodeError::invalid());
        }
        let field = parse_raw_field(source, cursor, limit)?;
        budget.field()?;
        cursor = field.end;
        match field.wire_type {
            3 => {
                let stack_depth =
                    u32::try_from(stack.len()).map_err(|_error| DecodeError::invalid())?;
                let child_depth = depth
                    .checked_add(stack_depth)
                    .ok_or_else(DecodeError::invalid)?;
                budget.observe_depth(child_depth)?;
                max_depth = max_depth.max(child_depth);
                let requested = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?;
                stack.try_reserve(1).map_err(|_error| {
                    DecodeError::resource(DecodeLimit::Allocations {
                        observed: requested,
                        maximum: stack.len(),
                    })
                })?;
                stack.push(field.number);
            },
            4 => {
                if field.number != group {
                    return Err(DecodeError::invalid());
                }
                stack.pop();
                if stack.is_empty() {
                    return Ok((cursor, max_depth));
                }
            },
            _ => {},
        }
    }
    Err(DecodeError::invalid())
}

fn next_raw_field(
    source: &[u8],
    offset: &mut usize,
    limit: usize,
    budget: &mut RewriteScanBudget,
    depth: u32,
) -> Result<Option<(RawField, u32)>, DecodeError> {
    if *offset >= limit {
        return Ok(None);
    }
    let mut field = parse_raw_field(source, *offset, limit)?;
    budget.field()?;
    let mut group_depth = 0;
    if field.wire_type == 3 {
        let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
        budget.observe_depth(child_depth)?;
        let (end, max_depth) = consume_raw_group(
            source,
            field.key_end,
            limit,
            field.number,
            budget,
            child_depth,
        )?;
        field.end = end;
        field.payload_end = end;
        group_depth = max_depth;
    } else if field.wire_type == 4 {
        return Err(DecodeError::invalid());
    }
    *offset = field.end;
    Ok(Some((field, group_depth)))
}

fn require_known_field(
    field: RawField,
    wire_type: u8,
    name: &'static str,
) -> Result<(), DecodeError> {
    if field.wire_type != wire_type {
        return Err(DecodeError::invalid());
    }
    if !field.key_canonical {
        return Err(DecodeError::noncanonical(name));
    }
    match wire_type {
        0 if !field.value_canonical => Err(DecodeError::noncanonical(name)),
        2 if !field.length_canonical => Err(DecodeError::noncanonical(name)),
        _ => Ok(()),
    }
}

fn validate_raw_date(
    source: &[u8],
    budget: &mut RewriteScanBudget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let mut offset = 0;
    let mut seconds = false;
    while let Some((field, _group_depth)) =
        next_raw_field(source, &mut offset, source.len(), budget, depth)?
    {
        if field.number == DATE_SECONDS_FIELD {
            if seconds {
                return Err(DecodeError::duplicate("TSP.Date.seconds"));
            }
            require_known_field(field, 1, "TSP.Date.seconds")?;
            seconds = true;
        }
    }
    if seconds {
        Ok(())
    } else {
        Err(DecodeError::missing("TSP.Date.seconds"))
    }
}

fn validate_raw_reference(
    source: &[u8],
    budget: &mut RewriteScanBudget,
    depth: u32,
) -> Result<RawReferenceFacts, DecodeError> {
    budget.message(source, depth)?;
    let before = budget.fields;
    let mut offset = 0;
    let mut identifier = None;
    let mut identifier_field = None;
    let mut deprecated_type = false;
    let mut deprecated_external = false;
    while let Some((field, _group_depth)) =
        next_raw_field(source, &mut offset, source.len(), budget, depth)?
    {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::duplicate("TSP.Reference.identifier"));
                }
                require_known_field(field, 0, "TSP.Reference.identifier")?;
                let (value, _length, _canonical) =
                    raw_varint(source, field.value_start, field.payload_end)?;
                if value == 0 {
                    return Err(DecodeError::invalid());
                }
                identifier = Some(value);
                identifier_field = Some(field);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type {
                    return Err(DecodeError::duplicate("TSP.Reference.deprecated_type"));
                }
                require_known_field(field, 0, "TSP.Reference.deprecated_type")?;
                let (value, _length, _canonical) =
                    raw_varint(source, field.value_start, field.payload_end)?;
                canonical_int32(value)?;
                deprecated_type = true;
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if deprecated_external {
                    return Err(DecodeError::duplicate(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                require_known_field(field, 0, "TSP.Reference.deprecated_is_external")?;
                let (value, _length, _canonical) =
                    raw_varint(source, field.value_start, field.payload_end)?;
                canonical_bool(value)?;
                deprecated_external = true;
            },
            _ => {},
        }
    }
    Ok(RawReferenceFacts {
        identifier: identifier.ok_or_else(|| DecodeError::missing("TSP.Reference.identifier"))?,
        identifier_field: identifier_field.ok_or_else(DecodeError::invalid)?,
        fields: budget
            .fields
            .checked_sub(before)
            .ok_or_else(DecodeError::invalid)?,
    })
}

fn validate_raw_uuid(
    source: &[u8],
    budget: &mut RewriteScanBudget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let mut offset = 0;
    let mut lower = false;
    let mut upper = false;
    while let Some((field, _group_depth)) =
        next_raw_field(source, &mut offset, source.len(), budget, depth)?
    {
        match field.number {
            UUID_LOWER_FIELD => {
                if lower {
                    return Err(DecodeError::duplicate("TSP.UUID.lower"));
                }
                require_known_field(field, 0, "TSP.UUID.lower")?;
                lower = true;
            },
            UUID_UPPER_FIELD => {
                if upper {
                    return Err(DecodeError::duplicate("TSP.UUID.upper"));
                }
                require_known_field(field, 0, "TSP.UUID.upper")?;
                upper = true;
            },
            _ => {},
        }
    }
    if lower && upper {
        Ok(())
    } else if !lower {
        Err(DecodeError::missing("TSP.UUID.lower"))
    } else {
        Err(DecodeError::missing("TSP.UUID.upper"))
    }
}

fn scan_comment_storage_raw(
    source: &[u8],
    options: DecodeOptions,
    target_ordinal: Option<usize>,
) -> Result<RawScanSummary, DecodeError> {
    let mut budget = RewriteScanBudget::new(source, options, options.max_message_bytes)?;
    budget.message(source, 1)?;
    let mut offset = 0;
    let mut text_seen = false;
    let mut date_seen = false;
    let mut author_seen = false;
    let mut uuid_seen = false;
    let mut reply_ids = Vec::new();
    let mut target = None;
    let mut target_reference = None;
    let mut max_group_depth = 0;
    while let Some((field, group_depth)) =
        next_raw_field(source, &mut offset, source.len(), &mut budget, 1)?
    {
        max_group_depth = max_group_depth.max(group_depth);
        match field.number {
            TEXT_FIELD => {
                if text_seen {
                    return Err(DecodeError::duplicate("TSD.CommentStorageArchive.text"));
                }
                require_known_field(field, 2, "TSD.CommentStorageArchive.text")?;
                str::from_utf8(&source[field.payload_start..field.payload_end])
                    .map_err(|_error| DecodeError::utf8("TSD.CommentStorageArchive.text"))?;
                budget.text(field.payload_len()?)?;
                text_seen = true;
            },
            CREATION_DATE_FIELD => {
                if date_seen {
                    return Err(DecodeError::duplicate(
                        "TSD.CommentStorageArchive.creation_date",
                    ));
                }
                require_known_field(field, 2, "TSD.CommentStorageArchive.creation_date")?;
                validate_raw_date(
                    &source[field.payload_start..field.payload_end],
                    &mut budget,
                    2,
                )?;
                date_seen = true;
            },
            AUTHOR_FIELD => {
                if author_seen {
                    return Err(DecodeError::duplicate("TSD.CommentStorageArchive.author"));
                }
                require_known_field(field, 2, "TSD.CommentStorageArchive.author")?;
                let raw = &source[field.payload_start..field.payload_end];
                let _facts = validate_raw_reference(raw, &mut budget, 2)?;
                budget.reference(raw.len(), false)?;
                author_seen = true;
            },
            REPLIES_FIELD => {
                require_known_field(field, 2, "TSD.CommentStorageArchive.replies")?;
                let raw = &source[field.payload_start..field.payload_end];
                let facts = validate_raw_reference(raw, &mut budget, 2)?;
                let ordinal = reply_ids.len();
                let requested = ordinal.checked_add(1).ok_or_else(DecodeError::invalid)?;
                reply_ids.try_reserve_exact(1).map_err(|_error| {
                    DecodeError::resource(DecodeLimit::Allocations {
                        observed: requested,
                        maximum: ordinal,
                    })
                })?;
                reply_ids.push(facts.identifier);
                budget.reference(raw.len(), true)?;
                if target_ordinal == Some(ordinal) {
                    target = Some(field);
                    target_reference = Some(facts);
                }
            },
            STORAGE_UUID_FIELD => {
                if uuid_seen {
                    return Err(DecodeError::duplicate(
                        "TSD.CommentStorageArchive.storage_uuid",
                    ));
                }
                require_known_field(field, 2, "TSD.CommentStorageArchive.storage_uuid")?;
                validate_raw_uuid(
                    &source[field.payload_start..field.payload_end],
                    &mut budget,
                    2,
                )?;
                uuid_seen = true;
            },
            _ => {},
        }
    }
    Ok(RawScanSummary {
        report: budget.report(),
        reply_ids,
        target,
        target_reference,
        max_group_depth,
    })
}

fn canonical_reply_payload_len(identifier: u64) -> Result<usize, DecodeError> {
    1usize
        .checked_add(encoded_varint_len(identifier))
        .ok_or_else(DecodeError::invalid)
}

fn canonical_reply_field_len(payload_len: usize) -> Result<usize, DecodeError> {
    1usize
        .checked_add(encoded_varint_len(
            u64::try_from(payload_len).map_err(|_error| DecodeError::invalid())?,
        ))
        .and_then(|length| length.checked_add(payload_len))
        .ok_or_else(DecodeError::invalid)
}

fn measure_reply_rewrite(
    source: &[u8],
    source_report: DecodeReport,
    rewrite: CommentStorageReplyRewrite,
    target: Option<RawField>,
    target_reference: Option<RawReferenceFacts>,
) -> Result<(usize, DecodeReport), DecodeError> {
    let (
        output_bytes,
        result_fields,
        result_work,
        result_references,
        result_replies,
        result_reference_bytes,
    ) = match rewrite {
        CommentStorageReplyRewrite::Append { identifier } => {
            let payload_len = canonical_reply_payload_len(identifier)?;
            let field_len = canonical_reply_field_len(payload_len)?;
            (
                source
                    .len()
                    .checked_add(field_len)
                    .ok_or_else(DecodeError::invalid)?,
                source_report
                    .fields
                    .checked_add(2)
                    .ok_or_else(DecodeError::invalid)?,
                source_report
                    .work_bytes
                    .checked_add(field_len)
                    .and_then(|work| work.checked_add(payload_len))
                    .ok_or_else(DecodeError::invalid)?,
                source_report
                    .references
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?,
                source_report
                    .replies
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?,
                source_report
                    .reference_bytes
                    .checked_add(payload_len)
                    .ok_or_else(DecodeError::invalid)?,
            )
        },
        CommentStorageReplyRewrite::Replace {
            replacement_identifier,
            ..
        } => {
            let target = target.ok_or_else(DecodeError::invalid)?;
            let target_reference = target_reference.ok_or_else(DecodeError::invalid)?;
            let old_payload_len = target.payload_len()?;
            let new_identifier_len = canonical_reply_payload_len(replacement_identifier)?;
            let new_payload_len = old_payload_len
                .checked_sub(target_reference.identifier_field.raw_len()?)
                .and_then(|length| length.checked_add(new_identifier_len))
                .ok_or_else(DecodeError::invalid)?;
            let new_field_len = canonical_reply_field_len(new_payload_len)?;
            let output_bytes = source
                .len()
                .checked_sub(target.raw_len()?)
                .and_then(|length| length.checked_add(new_field_len))
                .ok_or_else(DecodeError::invalid)?;
            let result_work = source_report
                .work_bytes
                .checked_sub(target.raw_len()?)
                .and_then(|work| work.checked_sub(old_payload_len))
                .and_then(|work| work.checked_add(new_field_len))
                .and_then(|work| work.checked_add(new_payload_len))
                .ok_or_else(DecodeError::invalid)?;
            (
                output_bytes,
                source_report.fields,
                result_work,
                source_report.references,
                source_report.replies,
                source_report
                    .reference_bytes
                    .checked_sub(old_payload_len)
                    .and_then(|bytes| bytes.checked_add(new_payload_len))
                    .ok_or_else(DecodeError::invalid)?,
            )
        },
        CommentStorageReplyRewrite::Remove { .. } => {
            let target = target.ok_or_else(DecodeError::invalid)?;
            let target_reference = target_reference.ok_or_else(DecodeError::invalid)?;
            let removed_fields = target_reference
                .fields
                .checked_add(1)
                .ok_or_else(DecodeError::invalid)?;
            (
                source
                    .len()
                    .checked_sub(target.raw_len()?)
                    .ok_or_else(DecodeError::invalid)?,
                source_report
                    .fields
                    .checked_sub(removed_fields)
                    .ok_or_else(DecodeError::invalid)?,
                source_report
                    .work_bytes
                    .checked_sub(target.raw_len()?)
                    .and_then(|work| work.checked_sub(target.payload_len().ok()?))
                    .ok_or_else(DecodeError::invalid)?,
                source_report
                    .references
                    .checked_sub(1)
                    .ok_or_else(DecodeError::invalid)?,
                source_report
                    .replies
                    .checked_sub(1)
                    .ok_or_else(DecodeError::invalid)?,
                source_report
                    .reference_bytes
                    .checked_sub(target.payload_len()?)
                    .ok_or_else(DecodeError::invalid)?,
            )
        },
    };
    let max_depth = source_report.max_depth.max(
        if matches!(rewrite, CommentStorageReplyRewrite::Append { .. }) {
            2
        } else {
            0
        },
    );
    Ok((
        output_bytes,
        DecodeReport {
            source_bytes: output_bytes,
            fields: result_fields,
            work_bytes: result_work,
            max_depth,
            references: result_references,
            replies: result_replies,
            reference_bytes: result_reference_bytes,
            text_bytes: source_report.text_bytes,
        },
    ))
}

fn emit_reply_rewrite(
    output: &mut Vec<u8>,
    source: &[u8],
    rewrite: CommentStorageReplyRewrite,
    target: Option<RawField>,
    target_reference: Option<RawReferenceFacts>,
    expected_length: usize,
) -> Result<(), DecodeError> {
    match rewrite {
        CommentStorageReplyRewrite::Append { identifier } => {
            output.extend_from_slice(source);
            append_reply_field(output, identifier)?;
        },
        CommentStorageReplyRewrite::Replace {
            replacement_identifier,
            ..
        } => {
            let target = target.ok_or_else(DecodeError::invalid)?;
            let reference = target_reference.ok_or_else(DecodeError::invalid)?;
            if reference.identifier == replacement_identifier {
                output.extend_from_slice(source);
            } else {
                output.extend_from_slice(&source[..target.start]);
                append_replaced_reply_field(
                    output,
                    source,
                    target,
                    reference,
                    replacement_identifier,
                )?;
                output.extend_from_slice(&source[target.end..]);
            }
        },
        CommentStorageReplyRewrite::Remove { .. } => {
            let target = target.ok_or_else(DecodeError::invalid)?;
            output.extend_from_slice(&source[..target.start]);
            output.extend_from_slice(&source[target.end..]);
        },
    }
    if output.len() == expected_length {
        Ok(())
    } else {
        Err(DecodeError::invalid())
    }
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn append_reply_field(output: &mut Vec<u8>, identifier: u64) -> Result<(), DecodeError> {
    let payload_len = canonical_reply_payload_len(identifier)?;
    output.push(0x22);
    append_varint(
        output,
        u64::try_from(payload_len).map_err(|_error| DecodeError::invalid())?,
    );
    output.push(0x08);
    append_varint(output, identifier);
    Ok(())
}

fn append_replaced_reply_field(
    output: &mut Vec<u8>,
    source: &[u8],
    target: RawField,
    reference: RawReferenceFacts,
    replacement_identifier: u64,
) -> Result<(), DecodeError> {
    let old_payload_len = target.payload_len()?;
    let new_identifier_len = canonical_reply_payload_len(replacement_identifier)?;
    let new_payload_len = old_payload_len
        .checked_sub(reference.identifier_field.raw_len()?)
        .and_then(|length| length.checked_add(new_identifier_len))
        .ok_or_else(DecodeError::invalid)?;
    output.push(0x22);
    append_varint(
        output,
        u64::try_from(new_payload_len).map_err(|_error| DecodeError::invalid())?,
    );
    let identifier_start = target
        .payload_start
        .checked_add(reference.identifier_field.start)
        .ok_or_else(DecodeError::invalid)?;
    let identifier_end = target
        .payload_start
        .checked_add(reference.identifier_field.end)
        .ok_or_else(DecodeError::invalid)?;
    output.extend_from_slice(&source[target.payload_start..identifier_start]);
    output.push(0x08);
    append_varint(output, replacement_identifier);
    output.extend_from_slice(&source[identifier_end..target.payload_end]);
    Ok(())
}

fn reply_ids_match(candidate: &[u64], source: &[u64], rewrite: CommentStorageReplyRewrite) -> bool {
    let expected_len = match rewrite {
        CommentStorageReplyRewrite::Append { .. } => source.len().checked_add(1),
        CommentStorageReplyRewrite::Replace { .. } => Some(source.len()),
        CommentStorageReplyRewrite::Remove { .. } => source.len().checked_sub(1),
    };
    let Some(expected_len) = expected_len else {
        return false;
    };
    if candidate.len() != expected_len {
        return false;
    }
    match rewrite {
        CommentStorageReplyRewrite::Append { identifier } => {
            candidate[..source.len()] == source[..] && candidate.last().copied() == Some(identifier)
        },
        CommentStorageReplyRewrite::Replace {
            ordinal,
            replacement_identifier,
            ..
        } => candidate.iter().enumerate().all(|(index, value)| {
            if index == ordinal {
                *value == replacement_identifier
            } else {
                source.get(index).copied() == Some(*value)
            }
        }),
        CommentStorageReplyRewrite::Remove { ordinal, .. } => {
            candidate.iter().enumerate().all(|(index, value)| {
                let Some(source_index) = (if index >= ordinal {
                    index.checked_add(1)
                } else {
                    Some(index)
                }) else {
                    return false;
                };
                source.get(source_index).copied() == Some(*value)
            })
        },
    }
}

fn check_rewrite_limits(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    if requirements.input_bytes > limits.input_bytes {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: requirements.input_bytes,
            maximum: limits.input_bytes,
        }));
    }
    if requirements.output_bytes > limits.output_bytes {
        return Err(DecodeError::resource(DecodeLimit::OutputBytes {
            observed: requirements.output_bytes,
            maximum: limits.output_bytes,
        }));
    }
    if requirements.fields > limits.fields {
        return Err(DecodeError::resource(DecodeLimit::Fields {
            observed: requirements.fields,
            maximum: limits.fields,
        }));
    }
    if requirements.work_bytes > limits.work_bytes {
        return Err(DecodeError::resource(DecodeLimit::Work {
            observed: requirements.work_bytes,
            maximum: limits.work_bytes,
        }));
    }
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::resource(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    if requirements.references > limits.references {
        return Err(DecodeError::resource(DecodeLimit::References {
            observed: requirements.references,
            maximum: limits.references,
        }));
    }
    if requirements.replies > limits.replies {
        return Err(DecodeError::resource(DecodeLimit::Replies {
            observed: requirements.replies,
            maximum: limits.replies,
        }));
    }
    if requirements.reference_bytes > limits.reference_bytes {
        return Err(DecodeError::resource(DecodeLimit::ReferenceBytes {
            observed: requirements.reference_bytes,
            maximum: limits.reference_bytes,
        }));
    }
    if requirements.allocations > limits.allocations {
        return Err(DecodeError::resource(DecodeLimit::Allocations {
            observed: requirements.allocations,
            maximum: limits.allocations,
        }));
    }
    if requirements.scratch_bytes > limits.scratch_bytes {
        return Err(DecodeError::resource(DecodeLimit::Scratch {
            observed: requirements.scratch_bytes,
            maximum: limits.scratch_bytes,
        }));
    }
    if requirements.retained_bytes > limits.retained_bytes {
        return Err(DecodeError::resource(DecodeLimit::Retained {
            observed: requirements.retained_bytes,
            maximum: limits.retained_bytes,
        }));
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "Focused wire fixtures intentionally use explicit test assertions."
)]
mod tests {
    use super::*;

    fn varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = u8::try_from(value & 0x7f).unwrap();
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

    fn key(output: &mut Vec<u8>, number: u32, wire_type: u8) {
        varint(output, (u64::from(number) << 3) | u64::from(wire_type));
    }

    fn field_varint(output: &mut Vec<u8>, number: u32, value: u64) {
        key(output, number, 0);
        varint(output, value);
    }

    fn field_bytes(output: &mut Vec<u8>, number: u32, value: &[u8]) {
        key(output, number, 2);
        varint(output, u64::try_from(value.len()).unwrap());
        output.extend_from_slice(value);
    }

    fn date(bits: u64) -> Vec<u8> {
        let mut output = Vec::new();
        key(&mut output, DATE_SECONDS_FIELD, 1);
        output.extend_from_slice(&bits.to_le_bytes());
        output
    }

    fn reference(
        identifier: u64,
        deprecated_type: Option<i32>,
        deprecated_is_external: Option<bool>,
    ) -> Vec<u8> {
        let mut output = Vec::new();
        field_varint(&mut output, REFERENCE_IDENTIFIER_FIELD, identifier);
        if let Some(value) = deprecated_type {
            field_varint(
                &mut output,
                REFERENCE_DEPRECATED_TYPE_FIELD,
                u64::from_ne_bytes(i64::from(value).to_ne_bytes()),
            );
        }
        if let Some(value) = deprecated_is_external {
            field_varint(
                &mut output,
                REFERENCE_DEPRECATED_EXTERNAL_FIELD,
                u64::from(value),
            );
        }
        output
    }

    fn uuid(lower: u64, upper: u64) -> Vec<u8> {
        let mut output = Vec::new();
        field_varint(&mut output, UUID_LOWER_FIELD, lower);
        field_varint(&mut output, UUID_UPPER_FIELD, upper);
        output
    }

    fn fixture() -> Vec<u8> {
        let mut output = Vec::new();
        field_bytes(&mut output, TEXT_FIELD, b"comment");
        field_bytes(
            &mut output,
            CREATION_DATE_FIELD,
            &date(0x8000_0000_0000_0000),
        );
        field_bytes(
            &mut output,
            AUTHOR_FIELD,
            &reference(11, Some(-7), Some(true)),
        );
        field_bytes(&mut output, REPLIES_FIELD, &reference(7, None, None));
        field_bytes(
            &mut output,
            REPLIES_FIELD,
            &reference(8, Some(4), Some(false)),
        );
        field_bytes(
            &mut output,
            STORAGE_UUID_FIELD,
            &uuid(0x0102_0304_0506_0708, 0x1112_1314_1516_1718),
        );
        output
    }

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            128,
            source.len().saturating_mul(32).max(1),
            8,
            32,
            4096,
        )
    }

    #[derive(Default)]
    struct Replies {
        identifiers: Vec<u64>,
        raw: Vec<Vec<u8>>,
    }

    impl CommentStorageVisitor for Replies {
        fn visit_reply(&mut self, reply: ReferenceRecord<'_>) -> Result<(), DecodeError> {
            self.identifiers.push(reply.identifier());
            self.raw.push(reply.raw().to_vec());
            Ok(())
        }
    }

    #[test]
    fn borrowed_projection_streams_replies_and_preserves_scalar_bits() {
        let source = fixture();
        let before = source.clone();
        let mut replies = Replies::default();
        let (snapshot, report) =
            decode_comment_storage_archive_with_visitor(&source, options(&source), &mut replies)
                .unwrap();

        assert_eq!(source, before);
        assert_eq!(snapshot.text(), Some("comment"));
        assert_eq!(
            snapshot.creation_date().map(DateSnapshot::seconds_bits),
            Some(0x8000_0000_0000_0000)
        );
        assert_eq!(
            snapshot.author().map(ReferenceSnapshot::identifier),
            Some(11)
        );
        assert_eq!(
            snapshot
                .author()
                .and_then(ReferenceSnapshot::deprecated_type),
            Some(-7)
        );
        assert_eq!(
            snapshot
                .author()
                .and_then(ReferenceSnapshot::deprecated_is_external),
            Some(true)
        );
        assert_eq!(
            snapshot
                .storage_uuid()
                .map(|uuid| (uuid.lower(), uuid.upper())),
            Some((0x0102_0304_0506_0708, 0x1112_1314_1516_1718))
        );
        assert_eq!(replies.identifiers, [7, 8]);
        assert_eq!(replies.raw.len(), 2);
        assert_eq!(report.references(), 3);
        assert_eq!(report.replies(), 2);
        assert_eq!(report.reply_references(), 2);
        assert_eq!(report.text_bytes(), 7);
        assert!(report.fields() >= 13);
        assert!(report.work_bytes() >= source.len() * 2);
        assert_eq!(replies.raw[0].as_slice(), &[8, 7]);
    }

    #[test]
    fn standalone_reference_charges_one_budget_unit_without_root_double_charge() {
        let source = reference(17, None, None);
        let rejected = decode_reference(
            &source,
            DecodeOptions::new(source.len(), 128, usize::MAX, 8, 0, 4096),
        )
        .unwrap_err();
        assert_eq!(
            rejected.resource_limit(),
            Some(DecodeLimit::References {
                observed: 1,
                maximum: 0,
            })
        );

        let accepted = decode_reference(
            &source,
            DecodeOptions::new(source.len(), 128, usize::MAX, 8, 1, 4096),
        )
        .unwrap();
        assert_eq!(accepted.identifier(), 17);

        let archive = fixture();
        let (_, report) =
            decode_comment_storage_archive_with_report(&archive, options(&archive)).unwrap();
        assert_eq!(report.references(), 3);
    }

    #[test]
    fn accepts_unknown_fields_and_well_formed_groups_without_retaining_them() {
        let mut source = fixture();
        let mut unknown = Vec::new();
        field_varint(&mut unknown, 99, 123);
        key(&mut unknown, 100, 3);
        field_varint(&mut unknown, 101, 8);
        key(&mut unknown, 100, 4);
        source.splice(0..0, unknown);
        assert!(decode_comment_storage_archive(&source, options(&source)).is_ok());
    }

    #[test]
    fn source_text_and_reply_payloads_borrow_input() {
        let source = fixture();
        struct PointerVisitor {
            reply_pointer: Option<*const u8>,
        }
        impl CommentStorageVisitor for PointerVisitor {
            fn visit_reply(&mut self, reply: ReferenceRecord<'_>) -> Result<(), DecodeError> {
                self.reply_pointer = Some(reply.raw().as_ptr());
                Ok(())
            }
        }
        let mut visitor = PointerVisitor {
            reply_pointer: None,
        };
        let (snapshot, _) =
            decode_comment_storage_archive_with_visitor(&source, options(&source), &mut visitor)
                .unwrap();
        let text = snapshot.text().unwrap();
        assert!(source.as_ptr_range().contains(&text.as_ptr()));
        assert!(
            source
                .as_ptr_range()
                .contains(&visitor.reply_pointer.unwrap())
        );
    }

    #[test]
    fn callback_error_is_returned_after_prior_replies() {
        struct Failing {
            seen: Vec<u64>,
        }
        impl CommentStorageVisitor for Failing {
            fn visit_reply(&mut self, reply: ReferenceRecord<'_>) -> Result<(), DecodeError> {
                self.seen.push(reply.identifier());
                if self.seen.len() == 2 {
                    return Err(DecodeError::invalid());
                }
                Ok(())
            }
        }
        let source = fixture();
        let mut visitor = Failing { seen: Vec::new() };
        assert!(
            decode_comment_storage_archive_with_visitor(&source, options(&source), &mut visitor)
                .is_err()
        );
        assert_eq!(visitor.seen, [7, 8]);
    }

    #[test]
    fn canonical_wire_and_schema_failures_are_rejected() {
        let source = fixture();
        let mut duplicate = source.clone();
        field_bytes(&mut duplicate, TEXT_FIELD, b"again");
        assert_eq!(
            decode_comment_storage_archive(&duplicate, options(&duplicate))
                .unwrap_err()
                .duplicate_singular_field(),
            Some("TSD.CommentStorageArchive.text")
        );

        let mut invalid_utf8 = Vec::new();
        field_bytes(&mut invalid_utf8, TEXT_FIELD, &[0xff]);
        assert_eq!(
            decode_comment_storage_archive(&invalid_utf8, options(&invalid_utf8))
                .unwrap_err()
                .invalid_utf8_field(),
            Some("TSD.CommentStorageArchive.text")
        );

        let mut missing_date = Vec::new();
        field_bytes(&mut missing_date, CREATION_DATE_FIELD, &[]);
        assert_eq!(
            decode_comment_storage_archive(&missing_date, options(&missing_date))
                .unwrap_err()
                .missing_required_field(),
            Some("TSP.Date.seconds")
        );

        let mut missing_reference = Vec::new();
        field_bytes(&mut missing_reference, AUTHOR_FIELD, &[]);
        assert_eq!(
            decode_comment_storage_archive(&missing_reference, options(&missing_reference))
                .unwrap_err()
                .missing_required_field(),
            Some("TSP.Reference.identifier")
        );

        let mut bad_bool = Vec::new();
        field_bytes(&mut bad_bool, AUTHOR_FIELD, &reference(1, None, Some(true)));
        let last = bad_bool.len() - 1;
        bad_bool[last] = 2;
        assert!(decode_comment_storage_archive(&bad_bool, options(&bad_bool)).is_err());

        let noncanonical_key = [0x88, 0x80, 0x00, 0x01];
        assert!(
            decode_comment_storage_archive(&noncanonical_key, options(&noncanonical_key)).is_err()
        );
        let noncanonical_value = [0x08, 0x81, 0x00];
        assert!(
            decode_comment_storage_archive(&noncanonical_value, options(&noncanonical_value))
                .is_err()
        );
        let noncanonical_known_length = [0x0a, 0x81, 0x00, b'a'];
        assert_eq!(
            decode_comment_storage_archive(
                &noncanonical_known_length,
                options(&noncanonical_known_length),
            )
            .unwrap_err()
            .noncanonical_reason(),
            Some("TSD.CommentStorageArchive.text")
        );
        let truncated = [0x0a, 0x04, b'a'];
        assert!(decode_comment_storage_archive(&truncated, options(&truncated)).is_err());
    }

    #[test]
    fn malformed_groups_and_wrong_wire_types_are_rejected() {
        let wrong_text_wire = [0x08, 0x01];
        assert!(
            decode_comment_storage_archive(&wrong_text_wire, options(&wrong_text_wire)).is_err()
        );
        let known_group = [0x0b, 0x0c];
        assert!(decode_comment_storage_archive(&known_group, options(&known_group)).is_err());
        let unclosed_group = [0x1b, 0x08, 0x01];
        assert!(decode_comment_storage_archive(&unclosed_group, options(&unclosed_group)).is_err());
        let mismatched_group = [0x1b, 0x24];
        assert!(
            decode_comment_storage_archive(&mismatched_group, options(&mismatched_group)).is_err()
        );
    }

    #[test]
    fn exact_resource_limits_are_enforced() {
        let source = fixture();
        let (_, report) =
            decode_comment_storage_archive_with_report(&source, options(&source)).unwrap();
        let bytes = decode_comment_storage_archive(
            &source,
            DecodeOptions::new(source.len() - 1, 128, usize::MAX, 8, 32, 4096),
        )
        .unwrap_err();
        assert_eq!(
            bytes.resource_limit(),
            Some(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1
            })
        );
        let fields = decode_comment_storage_archive(
            &source,
            DecodeOptions::new(source.len(), report.fields() - 1, usize::MAX, 8, 32, 4096),
        )
        .unwrap_err();
        assert!(matches!(
            fields.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        let work = decode_comment_storage_archive(
            &source,
            DecodeOptions::new(source.len(), 128, report.work_bytes() - 1, 8, 32, 4096),
        )
        .unwrap_err();
        assert!(matches!(
            work.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        let text = decode_comment_storage_archive(
            &source,
            DecodeOptions::new(
                source.len(),
                128,
                usize::MAX,
                8,
                32,
                report.text_bytes() - 1,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            text.resource_limit(),
            Some(DecodeLimit::Text { .. })
        ));
        let references = decode_comment_storage_archive(
            &source,
            DecodeOptions::new(
                source.len(),
                128,
                usize::MAX,
                8,
                report.references() - 1,
                4096,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            references.resource_limit(),
            Some(DecodeLimit::References { .. })
        ));
        let depth = decode_comment_storage_archive(
            &source,
            DecodeOptions::new(source.len(), 128, usize::MAX, 1, 32, 4096),
        )
        .unwrap_err();
        assert!(matches!(
            depth.resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));
    }

    fn reply_ids(source: &[u8]) -> Vec<u64> {
        let mut replies = Replies::default();
        decode_comment_storage_archive_with_visitor(source, options(source), &mut replies).unwrap();
        replies.identifiers
    }

    fn execute_reply_rewrite(
        source: &[u8],
        rewrite: CommentStorageReplyRewrite,
    ) -> (Vec<u8>, RewriteExecutionRequirements, RewriteReport) {
        let prepared =
            prepare_comment_storage_reply_rewrite(source, rewrite, options(source)).unwrap();
        let prepare_report = prepared.prepare_report();
        let requirements = prepared.execution_requirements();
        let output = prepared.execute(requirements.exact()).unwrap();
        let report = output.report();
        assert_eq!(report.source(), prepare_report);
        assert_eq!(report.input_bytes(), requirements.input_bytes);
        assert_eq!(report.output_bytes(), requirements.output_bytes);
        assert_eq!(report.fields(), requirements.fields);
        assert_eq!(report.work_bytes(), requirements.work_bytes);
        assert_eq!(report.max_depth(), requirements.max_depth);
        assert_eq!(report.references(), requirements.references);
        assert_eq!(report.replies(), requirements.replies);
        assert_eq!(report.reference_bytes(), requirements.reference_bytes);
        assert_eq!(report.allocations(), requirements.allocations);
        assert_eq!(report.scratch_bytes(), requirements.scratch_bytes);
        assert_eq!(report.retained_bytes(), requirements.retained_bytes);
        let bytes = output.into_bytes();
        (bytes, requirements, report)
    }

    #[test]
    fn prepared_reply_rewrite_appends_in_source_order_and_replays_report() {
        let source = fixture();
        let (output, requirements, _report) =
            execute_reply_rewrite(&source, CommentStorageReplyRewrite::append(19));
        assert_eq!(reply_ids(&output), [7, 8, 19]);
        assert_eq!(output.len(), requirements.output_bytes);
        assert_eq!(
            decode_comment_storage_archive(&output, options(&output))
                .unwrap()
                .text(),
            Some("comment")
        );
    }

    #[test]
    fn prepared_reply_rewrite_replace_and_remove_use_checked_ordinals() {
        let source = fixture();
        let (replaced, _, _) =
            execute_reply_rewrite(&source, CommentStorageReplyRewrite::replace(1, 8, 17));
        assert_eq!(reply_ids(&replaced), [7, 17]);
        let (restored, _, _) =
            execute_reply_rewrite(&replaced, CommentStorageReplyRewrite::replace(1, 17, 8));
        assert_eq!(restored, source);

        let (removed, _, _) =
            execute_reply_rewrite(&source, CommentStorageReplyRewrite::remove(0, 7));
        assert_eq!(reply_ids(&removed), [8]);

        let stale_ordinal = prepare_comment_storage_reply_rewrite(
            &source,
            CommentStorageReplyRewrite::replace(2, 8, 17),
            options(&source),
        )
        .unwrap_err();
        assert!(stale_ordinal.resource_limit().is_none());
        let stale_identifier = prepare_comment_storage_reply_rewrite(
            &source,
            CommentStorageReplyRewrite::remove(1, 7),
            options(&source),
        )
        .unwrap_err();
        assert!(stale_identifier.resource_limit().is_none());
    }

    #[test]
    fn prepared_reply_rewrite_rejects_duplicate_identities_and_preserves_noop_bytes() {
        let source = fixture();
        assert!(
            prepare_comment_storage_reply_rewrite(
                &source,
                CommentStorageReplyRewrite::append(7),
                options(&source),
            )
            .is_err()
        );
        assert!(
            prepare_comment_storage_reply_rewrite(
                &source,
                CommentStorageReplyRewrite::replace(0, 7, 8),
                options(&source),
            )
            .is_err()
        );
        let (noop, _, _) =
            execute_reply_rewrite(&source, CommentStorageReplyRewrite::replace(0, 7, 7));
        assert_eq!(noop, source);

        let mut duplicate_source = Vec::new();
        field_bytes(
            &mut duplicate_source,
            REPLIES_FIELD,
            &reference(7, None, None),
        );
        field_bytes(
            &mut duplicate_source,
            REPLIES_FIELD,
            &reference(7, None, None),
        );
        assert!(
            prepare_comment_storage_reply_rewrite(
                &duplicate_source,
                CommentStorageReplyRewrite::remove(0, 7),
                options(&duplicate_source),
            )
            .is_err()
        );
    }

    #[test]
    fn append_remove_is_exact_and_remove_only_reply_reports_actual_depth() {
        let source = fixture();
        let (appended, _, _) =
            execute_reply_rewrite(&source, CommentStorageReplyRewrite::append(19));
        let (restored, _, _) =
            execute_reply_rewrite(&appended, CommentStorageReplyRewrite::remove(2, 19));
        assert_eq!(restored, source);

        let mut only_reply = Vec::new();
        field_bytes(&mut only_reply, TEXT_FIELD, b"root");
        field_bytes(&mut only_reply, REPLIES_FIELD, &reference(31, None, None));
        let prepared = prepare_comment_storage_reply_rewrite(
            &only_reply,
            CommentStorageReplyRewrite::remove(0, 31),
            options(&only_reply),
        )
        .unwrap();
        let requirements = prepared.execution_requirements();
        let output = prepared.execute(requirements.exact()).unwrap();
        assert_eq!(output.report().result().max_depth(), 1);
        assert!(reply_ids(output.bytes()).is_empty());
    }

    #[test]
    fn reply_ordinals_are_independent_of_duplicate_text() {
        let mut source = Vec::new();
        field_bytes(&mut source, TEXT_FIELD, b"same");
        field_bytes(&mut source, REPLIES_FIELD, &reference(31, None, None));
        field_bytes(&mut source, REPLIES_FIELD, &reference(32, None, None));
        field_bytes(&mut source, REPLIES_FIELD, &reference(33, None, None));
        let (output, _, _) =
            execute_reply_rewrite(&source, CommentStorageReplyRewrite::replace(1, 32, 99));
        assert_eq!(reply_ids(&output), [31, 99, 33]);
    }

    #[test]
    fn zero_and_malformed_selected_reply_identifiers_fail_before_execution() {
        let source = fixture();
        for rewrite in [
            CommentStorageReplyRewrite::append(0),
            CommentStorageReplyRewrite::replace(0, 0, 7),
            CommentStorageReplyRewrite::replace(0, 7, 0),
            CommentStorageReplyRewrite::remove(0, 0),
        ] {
            assert!(
                prepare_comment_storage_reply_rewrite(&source, rewrite, options(&source)).is_err()
            );
        }

        let mut missing_identifier = Vec::new();
        field_bytes(&mut missing_identifier, REPLIES_FIELD, &[]);
        assert_eq!(
            prepare_comment_storage_reply_rewrite(
                &missing_identifier,
                CommentStorageReplyRewrite::remove(0, 1),
                options(&missing_identifier),
            )
            .unwrap_err()
            .missing_required_field(),
            Some("TSP.Reference.identifier")
        );

        let mut duplicate_identifier = reference(7, None, None);
        field_varint(&mut duplicate_identifier, REFERENCE_IDENTIFIER_FIELD, 8);
        let mut duplicate_root = Vec::new();
        field_bytes(&mut duplicate_root, REPLIES_FIELD, &duplicate_identifier);
        assert_eq!(
            prepare_comment_storage_reply_rewrite(
                &duplicate_root,
                CommentStorageReplyRewrite::remove(0, 7),
                options(&duplicate_root),
            )
            .unwrap_err()
            .duplicate_singular_field(),
            Some("TSP.Reference.identifier")
        );

        let mut wrong_wire_reference = Vec::new();
        key(&mut wrong_wire_reference, REFERENCE_IDENTIFIER_FIELD, 1);
        wrong_wire_reference.extend_from_slice(&7u64.to_le_bytes());
        let mut wrong_wire_root = Vec::new();
        field_bytes(&mut wrong_wire_root, REPLIES_FIELD, &wrong_wire_reference);
        assert!(
            prepare_comment_storage_reply_rewrite(
                &wrong_wire_root,
                CommentStorageReplyRewrite::remove(0, 7),
                options(&wrong_wire_root),
            )
            .is_err()
        );
    }

    #[test]
    fn unknown_overlong_scalars_and_balanced_groups_remain_byte_exact() {
        let mut prefix = Vec::new();
        prefix.extend_from_slice(&[0x98, 0x86, 0x00, 0x01]);
        key(&mut prefix, 99, 0);
        prefix.extend_from_slice(&[0xfb, 0x00]);
        key(&mut prefix, 105, 2);
        prefix.extend_from_slice(&[0x81, 0x00, 0x5a]);
        key(&mut prefix, 100, 3);
        field_varint(&mut prefix, 101, 8);
        key(&mut prefix, 100, 4);

        let mut suffix = Vec::new();
        key(&mut suffix, 102, 0);
        suffix.extend_from_slice(&[0x81, 0x00]);
        key(&mut suffix, 103, 3);
        field_varint(&mut suffix, 104, 1);
        key(&mut suffix, 103, 4);

        let mut source = prefix.clone();
        source.extend_from_slice(&fixture());
        source.extend_from_slice(&suffix);
        let (output, _, _) =
            execute_reply_rewrite(&source, CommentStorageReplyRewrite::replace(1, 8, 17));
        assert!(output.starts_with(&prefix));
        assert!(output.ends_with(&suffix));
        assert_eq!(
            scan_comment_storage_raw(&output, options(&output), None)
                .unwrap()
                .reply_ids,
            [7, 17]
        );
        assert_eq!(&output[..prefix.len()], prefix.as_slice());
        assert_eq!(&output[output.len() - suffix.len()..], suffix.as_slice());
        assert!(decode_comment_storage_archive(&output, options(&output)).is_ok());
    }

    #[test]
    fn append_candidate_depth_is_rejected_during_prepare() {
        let source = [0xa0, 0x06, 0x01];
        let limited = DecodeOptions::new(source.len(), 32, 128, 1, 8, 64);
        let error = prepare_comment_storage_reply_rewrite(
            &source,
            CommentStorageReplyRewrite::append(1),
            limited,
        )
        .unwrap_err();
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 2,
                maximum: 1,
            })
        );
    }

    #[test]
    fn prepared_raw_ingress_enforces_the_buffa_hard_message_ceiling() {
        let source = [0xa0, 0x06, 0x01];
        let hard = usize::try_from(buffa::MAX_MESSAGE_BYTES).unwrap();
        let above = hard.checked_add(1).unwrap();
        let error = prepare_comment_storage_reply_rewrite(
            &source,
            CommentStorageReplyRewrite::append(1),
            DecodeOptions::new(above, 32, 128, 2, 8, 64),
        )
        .unwrap_err();
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Bytes {
                observed: above,
                maximum: hard,
            })
        );
    }

    #[test]
    fn deeply_nested_unknown_groups_hit_the_typed_nesting_limit() {
        let mut source = Vec::new();
        for number in 100..=106 {
            key(&mut source, number, 3);
        }
        for number in (100..=106).rev() {
            key(&mut source, number, 4);
        }
        let limited = DecodeOptions::new(source.len(), 128, source.len() * 32, 4, 32, 4096);
        let error = prepare_comment_storage_reply_rewrite(
            &source,
            CommentStorageReplyRewrite::append(1),
            limited,
        )
        .unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed,
                maximum
            }) if observed > maximum
        ));
    }

    #[test]
    fn every_prepared_execution_axis_accepts_exact_and_rejects_one_below() {
        let source = fixture();
        let rewrite = CommentStorageReplyRewrite::append(55);
        let prepared =
            prepare_comment_storage_reply_rewrite(&source, rewrite, options(&source)).unwrap();
        let requirements = prepared.execution_requirements();
        let output = prepared.execute(requirements.exact()).unwrap();
        assert_eq!(output.report().output_bytes(), requirements.output_bytes);

        let below = |limits: RewriteExecutionLimits| {
            prepare_comment_storage_reply_rewrite(&source, rewrite, options(&source))
                .unwrap()
                .execute(limits)
                .unwrap_err()
        };
        let assert_axis = |error: DecodeError, expected: DecodeLimit| {
            assert_eq!(error.resource_limit(), Some(expected));
        };
        assert_axis(
            below(
                requirements
                    .exact()
                    .with_input_bytes(requirements.input_bytes - 1),
            ),
            DecodeLimit::Bytes {
                observed: requirements.input_bytes,
                maximum: requirements.input_bytes - 1,
            },
        );
        assert_axis(
            below(
                requirements
                    .exact()
                    .with_output_bytes(requirements.output_bytes - 1),
            ),
            DecodeLimit::OutputBytes {
                observed: requirements.output_bytes,
                maximum: requirements.output_bytes - 1,
            },
        );
        assert_axis(
            below(requirements.exact().with_fields(requirements.fields - 1)),
            DecodeLimit::Fields {
                observed: requirements.fields,
                maximum: requirements.fields - 1,
            },
        );
        assert_axis(
            below(
                requirements
                    .exact()
                    .with_work_bytes(requirements.work_bytes - 1),
            ),
            DecodeLimit::Work {
                observed: requirements.work_bytes,
                maximum: requirements.work_bytes - 1,
            },
        );
        assert_axis(
            below(
                requirements
                    .exact()
                    .with_max_depth(requirements.max_depth - 1),
            ),
            DecodeLimit::Nesting {
                observed: requirements.max_depth,
                maximum: requirements.max_depth - 1,
            },
        );
        assert_axis(
            below(
                requirements
                    .exact()
                    .with_references(requirements.references - 1),
            ),
            DecodeLimit::References {
                observed: requirements.references,
                maximum: requirements.references - 1,
            },
        );
        assert_axis(
            below(requirements.exact().with_replies(requirements.replies - 1)),
            DecodeLimit::Replies {
                observed: requirements.replies,
                maximum: requirements.replies - 1,
            },
        );
        assert_axis(
            below(
                requirements
                    .exact()
                    .with_reference_bytes(requirements.reference_bytes - 1),
            ),
            DecodeLimit::ReferenceBytes {
                observed: requirements.reference_bytes,
                maximum: requirements.reference_bytes - 1,
            },
        );
        assert_axis(
            below(
                requirements
                    .exact()
                    .with_allocations(requirements.allocations - 1),
            ),
            DecodeLimit::Allocations {
                observed: requirements.allocations,
                maximum: requirements.allocations - 1,
            },
        );
        assert_axis(
            below(
                requirements
                    .exact()
                    .with_scratch_bytes(requirements.scratch_bytes - 1),
            ),
            DecodeLimit::Scratch {
                observed: requirements.scratch_bytes,
                maximum: requirements.scratch_bytes - 1,
            },
        );
        assert_axis(
            below(
                requirements
                    .exact()
                    .with_retained_bytes(requirements.retained_bytes - 1),
            ),
            DecodeLimit::Retained {
                observed: requirements.retained_bytes,
                maximum: requirements.retained_bytes - 1,
            },
        );
    }
}
