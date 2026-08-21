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

use core::{fmt, str};

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
        self.references
    }
}

/// Typed finite resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    Bytes { observed: usize, maximum: usize },
    References { observed: usize, maximum: usize },
    Text { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
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
            DecodeErrorKind::Resource(DecodeLimit::References { .. }) => {
                formatter.write_str("Numbers comment-storage reference limit exceeded")
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
                if author.is_some() {
                    return Err(DecodeError::duplicate("TSD.CommentStorageArchive.author"));
                }
                let raw = field.bytes()?;
                budget.reference(raw.len())?;
                let value = decode_reference(raw, budget, child_depth)?;
                author = Some(value);
                raw_author = Some(raw);
            },
            REPLIES_FIELD => {
                let raw = field.bytes()?;
                budget.reference(raw.len())?;
                let reference = decode_reference(raw, budget, child_depth)?;
                parity_reference(raw, reference, budget, child_depth)?;
                visitor.visit_reply(ReferenceRecord { raw, reference })?;
            },
            STORAGE_UUID_FIELD => {
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

fn decode_reference(
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
                if identifier.is_some() {
                    return Err(DecodeError::duplicate("TSP.Reference.identifier"));
                }
                identifier = Some(field.varint()?);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type.is_some() {
                    return Err(DecodeError::duplicate("TSP.Reference.deprecated_type"));
                }
                deprecated_type = Some(canonical_int32(field.varint()?)?);
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
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
                if lower.is_some() {
                    return Err(DecodeError::duplicate("TSP.UUID.lower"));
                }
                lower = Some(field.varint()?);
            },
            UUID_UPPER_FIELD => {
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
    let (tag, canonical) = take_varint(source)?;
    if !canonical {
        return Err(DecodeError::noncanonical("protobuf field key"));
    }
    let raw_number = tag >> 3;
    let number =
        u32::try_from(raw_number).map_err(|_error| buffa::DecodeError::InvalidFieldNumber)?;
    if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let wire_type = u8::try_from(tag & 7).map_err(|_error| DecodeError::invalid())?;
    let value = match wire_type {
        0 => {
            let (value, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            }
            Value::Varint(value)
        },
        1 => Value::Fixed64(u64::from_le_bytes(
            take_exact(source, 8)?
                .try_into()
                .map_err(|_error| DecodeError::invalid())?,
        )),
        2 => {
            let (length, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("length-delimited size"));
            }
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
            let consumed = index + 1;
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

    fn reference(&mut self, bytes: usize) -> Result<(), DecodeError> {
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
            reference_bytes: self.reference_bytes,
            text_bytes: self.text_bytes,
        }
    }
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
        assert_eq!(report.reply_references(), 3);
        assert_eq!(report.text_bytes(), 7);
        assert!(report.fields() >= 13);
        assert!(report.work_bytes() >= source.len() * 2);
        assert_eq!(replies.raw[0].as_slice(), &[8, 7]);
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
}
