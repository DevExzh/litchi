//! Strict borrowed decoding for the shared `TSWP.NumberAttachmentArchive`.
//!
//! The raw payload remains owned by the format package.  This module performs
//! canonical wire validation, forces only the selected private Buffa lazy view,
//! and publishes borrowed semantic fields.  It deliberately has no encoder.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The decoder keeps schema preflight beside its lazy projection."
)]

use std::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_number_attachment_generated::LitchiIwaProjection as projection;

const SUPER_FIELD: u32 = 1;
const NUMBER_FORMAT_FIELD: u32 = 2;
const STRING_VALUE_FIELD: u32 = 3;
const NUMBER_FORMAT_NAME_FIELD: u32 = 4;
const TEXTUAL_STRING_EQUIVALENT_FIELD: u32 = 1;
const TEXTUAL_KIND_FIELD: u32 = 2;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Finite resource policy for one number-attachment payload.
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

    /// Build a conservative source-sized policy.
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

/// Exact strict fields/work consumed by a successful decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
}

impl DecodeReport {
    /// Number of wire records visited, including unknown records and groups.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate strict scan plus lazy projection work.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Deepest selected message, with the root at depth one.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }
}

/// A finite decoder resource observation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecodeLimit {
    /// Source or configured message bytes exceeded the limit.
    Bytes { observed: usize, maximum: usize },
    /// Wire records exceeded the limit.
    Fields { observed: usize, maximum: usize },
    /// Aggregate strict/lazy work exceeded the limit.
    Work { observed: usize, maximum: usize },
    /// Nested message depth exceeded the limit.
    Nesting { observed: u32, maximum: u32 },
}

/// Strict preflight or lazy projection failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Resource(DecodeLimit),
    Duplicate(&'static str),
    Missing(&'static str),
    WrongWire(&'static str),
    NonCanonical(&'static str),
    InvalidUtf8(&'static str),
    Invalid(&'static str),
    Projection,
}

impl DecodeError {
    const fn resource(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Resource(limit),
        }
    }

    const fn duplicate(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Duplicate(field),
        }
    }

    const fn missing(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Missing(field),
        }
    }

    const fn wrong_wire(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::WrongWire(field),
        }
    }

    const fn noncanonical(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(field),
        }
    }

    const fn invalid_utf8(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::InvalidUtf8(field),
        }
    }

    const fn invalid(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Invalid(field),
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    /// Return the finite resource observation, if one occurred.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Resource(limit) => Some(limit),
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(f),
            DecodeErrorKind::Resource(DecodeLimit::Bytes { observed, maximum }) => {
                write!(f, "number attachment bytes {observed} exceed {maximum}")
            },
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum }) => {
                write!(f, "number attachment fields {observed} exceed {maximum}")
            },
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum }) => {
                write!(f, "number attachment work {observed} exceeds {maximum}")
            },
            DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }) => {
                write!(f, "number attachment nesting {observed} exceeds {maximum}")
            },
            DecodeErrorKind::Duplicate(field) => {
                write!(f, "duplicate number attachment field {field}")
            },
            DecodeErrorKind::Missing(field) => write!(f, "missing number attachment field {field}"),
            DecodeErrorKind::WrongWire(field) => {
                write!(f, "wrong wire type for number attachment field {field}")
            },
            DecodeErrorKind::NonCanonical(field) => {
                write!(f, "noncanonical number attachment field {field}")
            },
            DecodeErrorKind::InvalidUtf8(field) => {
                write!(f, "invalid UTF-8 in number attachment field {field}")
            },
            DecodeErrorKind::Invalid(reason) => f.write_str(reason),
            DecodeErrorKind::Projection => {
                f.write_str("number attachment Buffa projection disagrees")
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

/// Borrowed semantic fields from one native number attachment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NumberAttachmentSnapshot<'source> {
    string_equivalent: Option<&'source str>,
    kind: Option<i32>,
    number_format: Option<u32>,
    string_value: Option<&'source str>,
    number_format_name: Option<&'source str>,
    raw: &'source [u8],
}

impl<'source> NumberAttachmentSnapshot<'source> {
    /// Borrow the optional textual equivalent.
    #[must_use]
    pub const fn string_equivalent(self) -> Option<&'source str> {
        self.string_equivalent
    }

    /// Return the optional native kind, preserving unknown values.
    #[must_use]
    pub const fn kind(self) -> Option<i32> {
        self.kind
    }

    /// Return the optional native number-format selector.
    #[must_use]
    pub const fn number_format(self) -> Option<u32> {
        self.number_format
    }

    /// Borrow the optional rendered string value.
    #[must_use]
    pub const fn string_value(self) -> Option<&'source str> {
        self.string_value
    }

    /// Borrow the optional format name.
    #[must_use]
    pub const fn number_format_name(self) -> Option<&'source str> {
        self.number_format_name
    }

    /// Return the exact source bytes validated by this decoder.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

/// Decode one complete `TSWP.NumberAttachmentArchive` without owning text.
pub fn decode_number_attachment<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<NumberAttachmentSnapshot<'source>, DecodeError> {
    Ok(decode_number_attachment_with_report(source, options)?.0)
}

/// Decode one attachment and report the bounded work used by the operation.
pub fn decode_number_attachment_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(NumberAttachmentSnapshot<'source>, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight(source, options, &mut budget)?;
    let view: projection::NumberAttachmentArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let projected_super = view
        .super_
        .get()
        .map_err(|_error| DecodeError::projection())?
        .ok_or_else(|| DecodeError::missing("TSWP.NumberAttachmentArchive.super"))?;
    let projected = NumberAttachmentSnapshot {
        string_equivalent: projected_super.string_equivalent,
        kind: projected_super.kind,
        number_format: view.number_format,
        string_value: view.string_value,
        number_format_name: view.number_format_name,
        raw: source,
    };
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok((strict, budget.finish()))
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let max_buffa = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_error| DecodeError::from(buffa::DecodeError::MessageTooLarge))?;
    if options.max_message_bytes == 0 || options.max_message_bytes > max_buffa {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: options.max_message_bytes,
            maximum: max_buffa,
        }));
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: options.max_message_bytes,
        }));
    }
    if options.max_fields == 0 {
        return Err(DecodeError::resource(DecodeLimit::Fields {
            observed: 0,
            maximum: options.max_fields,
        }));
    }
    if options.max_work_bytes == 0 {
        return Err(DecodeError::resource(DecodeLimit::Work {
            observed: 0,
            maximum: options.max_work_bytes,
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
    max_depth: u32,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
            max_depth: 0,
        }
    }

    fn field(&mut self) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(1);
        if observed > self.max_fields {
            return Err(DecodeError::resource(DecodeLimit::Fields {
                observed,
                maximum: self.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn message(&mut self, bytes: usize, depth: u32) -> Result<(), DecodeError> {
        self.max_depth = self.max_depth.max(depth);
        let observed = self.work_bytes.saturating_add(bytes.saturating_mul(2));
        if observed > self.max_work_bytes {
            return Err(DecodeError::resource(DecodeLimit::Work {
                observed,
                maximum: self.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn depth(&mut self, depth: u32) {
        self.max_depth = self.max_depth.max(depth);
    }

    fn finish(self) -> DecodeReport {
        DecodeReport {
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
        }
    }
}

fn preflight<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<NumberAttachmentSnapshot<'source>, DecodeError> {
    budget.message(source.len(), 1)?;
    let mut remaining = source;
    let mut super_payload = None;
    let mut number_format = None;
    let mut string_value = None;
    let mut number_format_name = None;
    while let Some(field) = next_field(&mut remaining, options, budget, 1)? {
        match field.number {
            SUPER_FIELD => {
                if super_payload.is_some() {
                    return Err(DecodeError::duplicate("super"));
                }
                super_payload = Some(field.bytes("super")?);
            },
            NUMBER_FORMAT_FIELD => {
                if number_format.is_some() {
                    return Err(DecodeError::duplicate("number_format"));
                }
                number_format = Some(field.u32("number_format")?);
            },
            STRING_VALUE_FIELD => {
                if string_value.is_some() {
                    return Err(DecodeError::duplicate("string_value"));
                }
                string_value = Some(field.string("string_value")?);
            },
            NUMBER_FORMAT_NAME_FIELD => {
                if number_format_name.is_some() {
                    return Err(DecodeError::duplicate("number_format_name"));
                }
                number_format_name = Some(field.string("number_format_name")?);
            },
            _ => {},
        }
    }
    let super_payload = super_payload.ok_or_else(|| DecodeError::missing("super"))?;
    let (string_equivalent, kind) = preflight_textual(super_payload, options, budget, 2)?;
    let kind = kind.ok_or_else(|| DecodeError::missing("TSWP.TextualAttachmentArchive.kind"))?;
    Ok(NumberAttachmentSnapshot {
        string_equivalent,
        kind: Some(kind),
        number_format,
        string_value,
        number_format_name,
        raw: source,
    })
}

fn preflight_textual<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<(Option<&'source str>, Option<i32>), DecodeError> {
    if depth > options.recursion_limit {
        return Err(DecodeError::resource(DecodeLimit::Nesting {
            observed: depth,
            maximum: options.recursion_limit,
        }));
    }
    budget.message(source.len(), depth)?;
    let mut remaining = source;
    let mut string_equivalent = None;
    let mut kind = None;
    while let Some(field) = next_field(&mut remaining, options, budget, depth)? {
        match field.number {
            TEXTUAL_STRING_EQUIVALENT_FIELD => {
                if string_equivalent.is_some() {
                    return Err(DecodeError::duplicate("string_equivalent"));
                }
                string_equivalent = Some(field.string("string_equivalent")?);
            },
            TEXTUAL_KIND_FIELD => {
                if kind.is_some() {
                    return Err(DecodeError::duplicate("kind"));
                }
                kind = Some(field.i32("kind")?);
            },
            _ => {},
        }
    }
    Ok((string_equivalent, kind))
}

#[derive(Debug, Clone, Copy)]
enum Value<'source> {
    Varint(u64),
    Fixed64,
    Bytes(&'source [u8]),
    Group,
    Fixed32,
}

#[derive(Debug, Clone, Copy)]
struct Field<'source> {
    number: u32,
    value: Value<'source>,
}

impl<'source> Field<'source> {
    fn bytes(self, name: &'static str) -> Result<&'source [u8], DecodeError> {
        match self.value {
            Value::Bytes(bytes) => Ok(bytes),
            _ => Err(DecodeError::wrong_wire(name)),
        }
    }

    fn string(self, name: &'static str) -> Result<&'source str, DecodeError> {
        str::from_utf8(self.bytes(name)?).map_err(|_error| DecodeError::invalid_utf8(name))
    }

    fn u32(self, name: &'static str) -> Result<u32, DecodeError> {
        let value = self.varint(name)?;
        u32::try_from(value).map_err(|_error| DecodeError::noncanonical(name))
    }

    fn i32(self, name: &'static str) -> Result<i32, DecodeError> {
        let value = self.varint(name)?;
        if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
            return Err(DecodeError::noncanonical(name));
        }
        Ok(value as i32)
    }

    fn varint(self, name: &'static str) -> Result<u64, DecodeError> {
        match self.value {
            Value::Varint(value) => Ok(value),
            _ => Err(DecodeError::wrong_wire(name)),
        }
    }
}

fn next_field<'source>(
    source: &mut &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<Field<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    let (tag, canonical) = take_varint(source)?;
    if !canonical {
        return Err(DecodeError::noncanonical("field key"));
    }
    budget.field()?;
    let number =
        u32::try_from(tag >> 3).map_err(|_error| DecodeError::invalid("invalid field number"))?;
    if number == 0 || number > 0x1fff_ffff {
        return Err(DecodeError::invalid("invalid field number"));
    }
    let wire_type = (tag & 7) as u8;
    let value = match wire_type {
        0 => {
            let (value, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("varint"));
            }
            Value::Varint(value)
        },
        1 => {
            take_exact(source, 8)?;
            Value::Fixed64
        },
        2 => {
            let (length, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("length"));
            }
            let length = usize::try_from(length)
                .map_err(|_error| DecodeError::invalid("length overflow"))?;
            Value::Bytes(take_exact(source, length)?)
        },
        3 => {
            let child_limit = options.recursion_limit.checked_sub(depth).ok_or_else(|| {
                DecodeError::resource(DecodeLimit::Nesting {
                    observed: depth.saturating_add(1),
                    maximum: options.recursion_limit,
                })
            })?;
            skip_group(
                source,
                number,
                child_limit,
                options,
                budget,
                depth.saturating_add(1),
            )?;
            Value::Group
        },
        4 => return Err(DecodeError::invalid("unexpected end group")),
        5 => {
            take_exact(source, 4)?;
            Value::Fixed32
        },
        _ => return Err(DecodeError::invalid("invalid wire type")),
    };
    Ok(Some(Field { number, value }))
}

fn skip_group(
    source: &mut &[u8],
    expected: u32,
    _remaining_depth: u32,
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    if depth > options.recursion_limit {
        return Err(DecodeError::resource(DecodeLimit::Nesting {
            observed: depth,
            maximum: options.recursion_limit,
        }));
    }
    budget.depth(depth);
    loop {
        if source.is_empty() {
            return Err(DecodeError::invalid("unterminated group"));
        }
        let (tag, canonical) = take_varint(source)?;
        if !canonical {
            return Err(DecodeError::noncanonical("group key"));
        }
        let number = u32::try_from(tag >> 3)
            .map_err(|_error| DecodeError::invalid("invalid group number"))?;
        let wire_type = (tag & 7) as u8;
        if wire_type == 4 {
            if number != expected {
                return Err(DecodeError::invalid("mismatched end group"));
            }
            budget.field()?;
            return Ok(());
        }
        // Reinsert the key in a tiny borrowed parse is impossible without
        // copying; parse the already consumed value directly and charge the
        // record here. Nested groups recurse through the same bounded path.
        budget.field()?;
        match wire_type {
            0 => {
                let (_, canonical) = take_varint(source)?;
                if !canonical {
                    return Err(DecodeError::noncanonical("group varint"));
                }
            },
            1 => {
                take_exact(source, 8)?;
            },
            2 => {
                let (length, canonical) = take_varint(source)?;
                if !canonical {
                    return Err(DecodeError::noncanonical("group length"));
                }
                let length = usize::try_from(length)
                    .map_err(|_error| DecodeError::invalid("group length overflow"))?;
                take_exact(source, length)?;
            },
            3 => {
                if depth >= options.recursion_limit {
                    return Err(DecodeError::resource(DecodeLimit::Nesting {
                        observed: depth + 1,
                        maximum: options.recursion_limit,
                    }));
                }
                skip_group(
                    source,
                    number,
                    options.recursion_limit - depth,
                    options,
                    budget,
                    depth + 1,
                )?;
            },
            5 => {
                take_exact(source, 4)?;
            },
            _ => return Err(DecodeError::invalid("invalid group wire type")),
        }
    }
}

fn take_varint(source: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original
            .get(index)
            .ok_or_else(|| DecodeError::invalid("truncated varint"))?;
        if index == 9 && byte > 1 {
            return Err(DecodeError::invalid("varint overflow"));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            *source = &original[consumed..];
            return Ok((value, canonical_varint_len(value) == consumed));
        }
    }
    Err(DecodeError::invalid("varint too long"))
}

fn canonical_varint_len(mut value: u64) -> usize {
    let mut length = 1;
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
        return Err(DecodeError::invalid("truncated field"));
    }
    let (selected, remaining) = source.split_at(length);
    *source = remaining;
    Ok(selected)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn varint(mut value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return output;
            }
        }
    }

    fn bytes(field: u32, value: &[u8]) -> Vec<u8> {
        let mut output = varint(u64::from(field) << 3 | 2);
        output.extend(varint(value.len() as u64));
        output.extend(value);
        output
    }

    fn scalar(field: u32, value: u64) -> Vec<u8> {
        let mut output = varint(u64::from(field) << 3);
        output.extend(varint(value));
        output
    }

    fn group(field: u32, nested: &[u8]) -> Vec<u8> {
        let mut output = varint(u64::from(field) << 3 | 3);
        output.extend(nested);
        output.extend(varint(u64::from(field) << 3 | 4));
        output
    }

    fn fixture() -> Vec<u8> {
        let mut textual = bytes(1, b"Page ");
        textual.extend(scalar(2, 1));
        let mut source = bytes(1, &textual);
        source.extend(scalar(2, 3));
        source.extend(bytes(3, b"12"));
        source.extend(bytes(4, b"decimal"));
        source
    }

    #[test]
    fn decodes_all_selected_fields_without_owning_them() {
        let source = fixture();
        let (snapshot, report) =
            decode_number_attachment_with_report(&source, DecodeOptions::for_source(&source))
                .expect("attachment should decode");
        assert_eq!(snapshot.string_equivalent(), Some("Page "));
        assert_eq!(snapshot.kind(), Some(1));
        assert_eq!(snapshot.number_format(), Some(3));
        assert_eq!(snapshot.string_value(), Some("12"));
        assert_eq!(snapshot.number_format_name(), Some("decimal"));
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(report.fields(), 6);
        assert!(report.work_bytes() >= source.len() * 2);
    }

    #[test]
    fn rejects_duplicates_noncanonical_text_and_tight_limits() {
        let mut duplicate = fixture();
        duplicate.extend(bytes(3, b"again"));
        assert!(
            decode_number_attachment(&duplicate, DecodeOptions::for_source(&duplicate)).is_err()
        );
        let invalid = bytes(3, &[0xff]);
        let mut source = bytes(1, &bytes(1, b""));
        source.extend(invalid);
        assert!(decode_number_attachment(&source, DecodeOptions::for_source(&source)).is_err());
        let source = fixture();
        let report =
            decode_number_attachment_with_report(&source, DecodeOptions::for_source(&source))
                .expect("measure attachment work")
                .1;
        let options = DecodeOptions::new(source.len(), report.fields(), report.work_bytes() - 1, 8);
        assert!(matches!(
            decode_number_attachment(&source, options),
            Err(error) if matches!(error.resource_limit(), Some(DecodeLimit::Work { .. }))
        ));
    }

    #[test]
    fn unknown_fields_are_retained_in_raw_source() {
        let mut source = fixture();
        source.extend(scalar(99, 1));
        let snapshot = decode_number_attachment(&source, DecodeOptions::for_source(&source))
            .expect("unknown field is source-owned");
        assert_eq!(snapshot.raw(), source.as_slice());
    }

    #[test]
    fn unknown_nested_groups_contribute_to_reported_depth() {
        let mut source = fixture();
        let nested = group(100, &scalar(101, 1));
        source.extend(group(99, &nested));
        let (_, report) =
            decode_number_attachment_with_report(&source, DecodeOptions::for_source(&source))
                .expect("unknown groups are skipped within the finite wire profile");
        assert_eq!(report.max_depth(), 3);
    }

    #[test]
    fn group_at_textual_message_depth_limit_is_rejected() {
        let mut textual = bytes(1, b"Page ");
        textual.extend(scalar(2, 1));
        textual.extend(group(99, &scalar(100, 1)));
        let source = bytes(1, &textual);
        let mut options = DecodeOptions::for_source(&source);
        options.recursion_limit = 3;
        let (_, report) = decode_number_attachment_with_report(&source, options)
            .expect("group exactly at depth limit");
        assert_eq!(report.max_depth(), 3);
        options.recursion_limit = 2;
        let error = decode_number_attachment_with_report(&source, options)
            .expect_err("direct group must not enter depth three with limit two");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 3,
                maximum: 2
            })
        ));
    }

    #[test]
    fn missing_textual_kind_is_rejected_before_projection() {
        let source = bytes(1, &bytes(1, b"Page "));
        assert!(matches!(
            decode_number_attachment(&source, DecodeOptions::for_source(&source)),
            Err(error) if error.to_string().contains("missing number attachment field")
        ));
    }
}
