//! Strict raw-preserving Buffa projection for a Pages footnote marker.
//!
//! The handwritten pass validates the two selected `TextualAttachmentArchive`
//! fields and every source field's wire framing before Buffa observes the
//! payload. Unknown fields and all source bytes remain owned by the caller;
//! this codec has no production encoding or rewriting path.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict semantic preflight intentionally precedes the low-level wire reader."
)]

use std::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_pages_footnote_marker_generated::LitchiIwaProjection as projection;

const TEXTUAL_STRING_EQUIVALENT_FIELD: u32 = 1;
const TEXTUAL_KIND_FIELD: u32 = 2;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Explicit finite resource policy for one Pages footnote marker payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build a finite bytes/fields/work/nesting policy.
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
        }
    }

    /// Build conservative finite limits from one known source length.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(4).max(1),
            bytes.saturating_mul(16).max(1),
            8,
        )
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Borrowed semantic facts from one Pages footnote marker.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextualAttachmentSnapshot<'source> {
    string_equivalent: Option<&'source str>,
    kind: Option<i32>,
    raw: &'source [u8],
}

impl<'source> TextualAttachmentSnapshot<'source> {
    /// Optional native string-equivalent, borrowed from the source payload.
    #[must_use]
    pub const fn string_equivalent(self) -> Option<&'source str> {
        self.string_equivalent
    }

    /// Optional native attachment kind, including unknown enum values.
    #[must_use]
    pub const fn kind(self) -> Option<i32> {
        self.kind
    }

    /// Return the exact caller-owned bytes that were preflighted.
    ///
    /// The projection never rewrites, normalizes, or stores unknown fields;
    /// callers can retain this slice as their lossless preservation value.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

/// Failure from strict Pages footnote-marker preflight or its Buffa
/// cross-check.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

/// Finite resource rejected by the Pages footnote-marker codec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Source or configured Buffa message bytes exceeded the ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Strict wire-field visits exceeded the ceiling.
    Fields { observed: usize, maximum: usize },
    /// Strict plus lazy-projection work exceeded the ceiling.
    Work { observed: usize, maximum: usize },
    /// Configured or traversed protobuf nesting exceeded the ceiling.
    Nesting { observed: u32, maximum: u32 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Resource(DecodeLimit),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    InvalidUtf8(&'static str),
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
    Projection,
}

impl DecodeError {
    const fn recursion_limit(observed: u32, maximum: u32) -> Self {
        Self::resource(DecodeLimit::Nesting { observed, maximum })
    }

    const fn resource(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Resource(limit),
        }
    }

    const fn duplicate_singular(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateSingular(field),
        }
    }

    const fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason),
        }
    }

    const fn invalid_utf8(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::InvalidUtf8(field),
        }
    }

    const fn field_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::FieldLimit { observed, maximum },
        }
    }

    const fn work_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::WorkLimit { observed, maximum },
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    /// Singular known field repeated in the source, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::InvalidUtf8(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Stable canonicality failure reason, when applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::InvalidUtf8(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Text field carrying invalid UTF-8, when applicable.
    #[must_use]
    pub const fn invalid_utf8_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::InvalidUtf8(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Observed/configured field counts for a field-limit failure.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::FieldLimit { observed, maximum } => Some((observed, maximum)),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::InvalidUtf8(_)
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Observed/configured work bytes for a work-limit failure.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::WorkLimit { observed, maximum } => Some((observed, maximum)),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::InvalidUtf8(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Return the typed finite resource failure, when applicable.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Resource(limit) => Some(limit),
            DecodeErrorKind::FieldLimit { observed, maximum } => {
                Some(DecodeLimit::Fields { observed, maximum })
            },
            DecodeErrorKind::WorkLimit { observed, maximum } => {
                Some(DecodeLimit::Work { observed, maximum })
            },
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::InvalidUtf8(_)
            | DecodeErrorKind::Projection => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Resource(DecodeLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "Pages footnote-marker projection byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "Pages footnote-marker projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "Pages footnote-marker projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Pages footnote-marker projection nesting limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::InvalidUtf8(field) => write!(formatter, "{field} is invalid UTF-8"),
            DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Pages footnote-marker projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Pages footnote-marker projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Projection => formatter.write_str(
                "Pages footnote-marker strict preflight disagrees with the Buffa projection",
            ),
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

/// Decode one complete `TSWP.TextualAttachmentArchive` marker payload.
pub fn decode_textual_attachment<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<TextualAttachmentSnapshot<'source>, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_textual(source, options, &mut budget)?;
    let view: projection::TextualAttachmentArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let projected = force_textual_projection(&view, source)?;
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let max_buffa_message_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
    if options.max_message_bytes > max_buffa_message_bytes {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: options.max_message_bytes,
            maximum: max_buffa_message_bytes,
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
    Ok(())
}

#[derive(Debug)]
struct Budget {
    fields: usize,
    work_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
        }
    }

    fn charge_field(&mut self) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(1);
        if observed > self.max_fields {
            return Err(DecodeError::field_limit(observed, self.max_fields));
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_message(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let strict_and_projection = bytes.saturating_mul(2);
        let observed = self.work_bytes.saturating_add(strict_and_projection);
        if observed > self.max_work_bytes {
            return Err(DecodeError::work_limit(observed, self.max_work_bytes));
        }
        self.work_bytes = observed;
        Ok(())
    }
}

fn preflight_textual<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<TextualAttachmentSnapshot<'source>, DecodeError> {
    budget.charge_message(source.len())?;
    let mut string_equivalent = None;
    let mut kind = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            TEXTUAL_STRING_EQUIVALENT_FIELD => {
                if string_equivalent.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.TextualAttachmentArchive.string_equivalent",
                    ));
                }
                string_equivalent =
                    Some(str::from_utf8(field.length_delimited()?).map_err(|_error| {
                        DecodeError::invalid_utf8("TSWP.TextualAttachmentArchive.string_equivalent")
                    })?);
            },
            TEXTUAL_KIND_FIELD => {
                if kind.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.TextualAttachmentArchive.kind",
                    ));
                }
                kind = Some(decode_int32(require_canonical_int32(field.varint()?)?));
            },
            _ => {},
        }
    }
    Ok(TextualAttachmentSnapshot {
        string_equivalent,
        kind,
        raw: source,
    })
}

fn force_textual_projection<'source>(
    view: &projection::TextualAttachmentArchiveLazyView<'source>,
    source: &'source [u8],
) -> Result<TextualAttachmentSnapshot<'source>, DecodeError> {
    Ok(TextualAttachmentSnapshot {
        string_equivalent: view.string_equivalent,
        kind: view.kind,
        raw: source,
    })
}

fn require_canonical_int32(value: u64) -> Result<u64, DecodeError> {
    if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical(
            "int32 scalar is not a sign-extended 32-bit value",
        ));
    }
    Ok(value)
}

#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    reason = "Strict preflight proved the u64 is a canonical sign-extended int32."
)]
fn decode_int32(value: u64) -> i32 {
    value as i32
}

#[derive(Clone, Copy, Debug)]
enum StrictValue<'source> {
    Varint(u64),
    Fixed64,
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32,
}

#[derive(Clone, Copy, Debug)]
struct StrictField<'source> {
    number: u32,
    wire_type: buffa::encoding::WireType,
    value: StrictValue<'source>,
}

impl<'source> StrictField<'source> {
    fn require_wire_type(self, expected: buffa::encoding::WireType) -> Result<(), DecodeError> {
        if self.wire_type != expected {
            return Err(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected: expected as u8,
                actual: self.wire_type as u8,
            }
            .into());
        }
        Ok(())
    }

    fn varint(self) -> Result<u64, DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::Varint)?;
        match self.value {
            StrictValue::Varint(value) => Ok(value),
            StrictValue::Fixed64
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group
            | StrictValue::Fixed32 => Err(DecodeError::projection()),
        }
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::LengthDelimited)?;
        match self.value {
            StrictValue::LengthDelimited(value) => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64
            | StrictValue::Group
            | StrictValue::Fixed32 => Err(DecodeError::projection()),
        }
    }
}

enum ParseItem<'source> {
    Field(StrictField<'source>),
    EndGroup(u32),
}

fn next_strict_field<'source>(
    source: &mut &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_strict_field(source, options.recursion_limit, budget)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(number)) => {
            Err(buffa::DecodeError::InvalidEndGroup(number).into())
        },
        None => Ok(None),
    }
}

fn parse_strict_field<'source>(
    source: &mut &'source [u8],
    recursion_limit: u32,
    budget: &mut Budget,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    let (encoded_tag, canonical_key) = take_varint(source)?;
    if !canonical_key {
        return Err(DecodeError::noncanonical("protobuf field key"));
    }
    budget.charge_field()?;
    let raw_tag =
        u32::try_from(encoded_tag).map_err(|_conversion| buffa::DecodeError::InvalidFieldNumber)?;
    let field_number = raw_tag >> 3;
    if field_number == 0 || field_number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let raw_wire_type = raw_tag & 7;
    let wire_type = buffa::encoding::WireType::from_u32(raw_wire_type)?;
    let value = match wire_type {
        buffa::encoding::WireType::Varint => {
            let (value, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            }
            StrictValue::Varint(value)
        },
        buffa::encoding::WireType::Fixed64 => {
            take_exact(source, 8)?;
            StrictValue::Fixed64
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (encoded_length, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("length-delimited size"));
            }
            let length = usize::try_from(encoded_length)
                .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
            StrictValue::LengthDelimited(take_exact(source, length)?)
        },
        buffa::encoding::WireType::StartGroup => {
            let child_limit = recursion_limit.checked_sub(1).ok_or_else(|| {
                DecodeError::recursion_limit(recursion_limit.saturating_add(1), recursion_limit)
            })?;
            skip_strict_group(source, field_number, child_limit, budget)?;
            StrictValue::Group
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            take_exact(source, 4)?;
            StrictValue::Fixed32
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw_wire_type).into()),
    };
    Ok(Some(ParseItem::Field(StrictField {
        number: field_number,
        wire_type,
        value,
    })))
}

fn skip_strict_group(
    source: &mut &[u8],
    expected_field_number: u32,
    recursion_limit: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    loop {
        match parse_strict_field(source, recursion_limit, budget)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected_field_number => return Ok(()),
            Some(ParseItem::EndGroup(number)) => {
                return Err(buffa::DecodeError::InvalidEndGroup(number).into());
            },
            None => return Err(buffa::DecodeError::UnexpectedEof.into()),
        }
    }
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
            return Ok((value, canonical_varint_len(value) == consumed));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
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
    source: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], DecodeError> {
    if source.len() < length {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    let (selected, remaining) = source.split_at(length);
    *source = remaining;
    Ok(selected)
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::shadow_unrelated,
    reason = "Focused codec tests use explicit fixture panic messages."
)]
mod tests {
    use prost::Message as _;

    use super::*;
    use crate::tswp;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
    }

    fn fixture() -> tswp::TextualAttachmentArchive {
        tswp::TextualAttachmentArchive {
            string_equivalent: Some("*".to_owned()),
            kind: Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32),
        }
    }

    #[test]
    fn canonical_fixture_matches_private_projection() -> Result<(), Box<dyn std::error::Error>> {
        let source = fixture().encode_to_vec();
        let snapshot = decode_textual_attachment(&source, options(&source))?;
        assert_eq!(snapshot.string_equivalent(), Some("*"));
        assert_eq!(snapshot.kind(), Some(2));
        assert_eq!(snapshot.raw(), source.as_slice());
        Ok(())
    }

    #[test]
    fn optional_fields_preserve_absence() -> Result<(), Box<dyn std::error::Error>> {
        let source = [];
        let snapshot = decode_textual_attachment(&source, options(&source))?;
        assert_eq!(snapshot.string_equivalent(), None);
        assert_eq!(snapshot.kind(), None);
        assert!(snapshot.raw().is_empty());
        Ok(())
    }

    #[test]
    fn unknown_fields_remain_opaque_but_strictly_framed() -> Result<(), Box<dyn std::error::Error>>
    {
        let mut source = fixture().encode_to_vec();
        let unknown = [0xa0, 0x06, 0x01, 0xaa, 0x06, 0x03, 0x6f, 0x70, 0x61];
        source.extend_from_slice(&unknown);
        let snapshot = decode_textual_attachment(&source, options(&source))?;
        assert_eq!(snapshot.kind(), Some(2));
        // The lazy projection intentionally drops unknown fields.  The
        // caller-owned source remains the only lossless representation, so
        // keep this assertion exact rather than merely checking that the
        // opaque suffix survived.
        assert_eq!(snapshot.raw(), source.as_slice());
        assert!(snapshot.raw().ends_with(&unknown));
        Ok(())
    }

    #[test]
    fn malformed_known_fields_are_rejected_before_lazy_projection() {
        let mut duplicate = fixture().encode_to_vec();
        duplicate.extend_from_slice(&[0x10, 0x01]);
        let error = decode_textual_attachment(&duplicate, options(&duplicate))
            .expect_err("duplicate marker kind");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSWP.TextualAttachmentArchive.kind")
        );

        let bad_utf8 = [0x0a, 0x01, 0xff];
        let error = decode_textual_attachment(&bad_utf8, options(&bad_utf8))
            .expect_err("invalid marker text");
        assert_eq!(
            error.invalid_utf8_field(),
            Some("TSWP.TextualAttachmentArchive.string_equivalent")
        );

        let malformed_unknown = [0xa0, 0x06, 0x80];
        assert!(
            decode_textual_attachment(&malformed_unknown, options(&malformed_unknown)).is_err()
        );
    }

    #[test]
    fn exact_limits_are_enforced() {
        let source = fixture().encode_to_vec();
        assert_eq!(
            decode_textual_attachment(
                &source,
                DecodeOptions::new(source.len() - 1, source.len() * 4, source.len() * 16, 8)
            )
            .expect_err("byte ceiling")
            .resource_limit(),
            Some(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1,
            })
        );
        assert!(matches!(
            decode_textual_attachment(
                &source,
                DecodeOptions::new(source.len(), 1, source.len() * 16, 8)
            )
            .expect_err("field ceiling")
            .resource_limit(),
            Some(DecodeLimit::Fields {
                observed: 2,
                maximum: 1,
            })
        ));
        assert!(matches!(
            decode_textual_attachment(
                &source,
                DecodeOptions::new(source.len(), source.len() * 4, source.len(), 8)
            )
            .expect_err("work ceiling")
            .resource_limit(),
            Some(DecodeLimit::Work {
                observed: _,
                maximum,
            }) if maximum == source.len()
        ));
        assert_eq!(
            decode_textual_attachment(
                &source,
                DecodeOptions::new(source.len(), source.len() * 4, source.len() * 16, 0)
            )
            .expect_err("nesting ceiling")
            .resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 0,
                maximum: MAX_RECURSION_LIMIT,
            })
        );
    }
}
