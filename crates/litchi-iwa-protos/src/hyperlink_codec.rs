//! Bounded raw codec for the complete `TSWP.HyperlinkFieldArchive` schema.
//!
//! The native schema has only two optional fields: a nested
//! `SmartFieldArchive` carrying the text-attribute UUID and the hyperlink URL.
//! Keeping this seam raw avoids materializing generated strings during object
//! validation.  Callers retain the original bytes when editing, so unknown
//! fields and their original order remain untouched by the codec.

use std::{fmt, str};

const SUPER_FIELD: u32 = 1;
const URL_FIELD: u32 = 2;
const UUID_FIELD: u32 = 1;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;

/// Finite resource policy for one hyperlink payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build an explicit bytes/fields/work/nesting policy.
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

    /// Build a finite profile for a source slice already owned by the caller.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(4).max(1),
            bytes.saturating_mul(4).max(1),
            8,
        )
    }
}

/// Borrowed semantic facts from one hyperlink payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HyperlinkSnapshot<'source> {
    text_attribute_uuid: Option<&'source str>,
    url_ref: Option<&'source str>,
}

impl<'source> HyperlinkSnapshot<'source> {
    /// Return the optional text-attribute UUID without allocating.
    #[must_use]
    pub const fn text_attribute_uuid(self) -> Option<&'source str> {
        self.text_attribute_uuid
    }

    /// Return the optional hyperlink target without allocating.
    #[must_use]
    pub const fn url_ref(self) -> Option<&'source str> {
        self.url_ref
    }
}

/// Resource class for a bounded raw decode failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecodeLimit {
    /// Source bytes exceeded the configured message ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Visited fields exceeded the configured field ceiling.
    Fields { observed: usize, maximum: usize },
    /// Wire work exceeded the configured work ceiling.
    Work { observed: usize, maximum: usize },
    /// Nested groups/messages exceeded the configured recursion ceiling.
    Nesting { observed: u32, maximum: u32 },
}

/// Strict hyperlink decode failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeError(DecodeErrorKind);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DecodeErrorKind {
    InvalidWire,
    Truncated,
    NonCanonical(&'static str),
    InvalidUtf8(&'static str),
    Duplicate(&'static str),
    UnexpectedEndGroup,
    MissingEndGroup,
    Limit(DecodeLimit),
}

impl DecodeError {
    const fn invalid_wire() -> Self {
        Self(DecodeErrorKind::InvalidWire)
    }

    const fn truncated() -> Self {
        Self(DecodeErrorKind::Truncated)
    }

    const fn noncanonical(reason: &'static str) -> Self {
        Self(DecodeErrorKind::NonCanonical(reason))
    }

    const fn invalid_utf8(field: &'static str) -> Self {
        Self(DecodeErrorKind::InvalidUtf8(field))
    }

    const fn duplicate(field: &'static str) -> Self {
        Self(DecodeErrorKind::Duplicate(field))
    }

    const fn unexpected_end_group() -> Self {
        Self(DecodeErrorKind::UnexpectedEndGroup)
    }

    const fn missing_end_group() -> Self {
        Self(DecodeErrorKind::MissingEndGroup)
    }

    const fn limit(limit: DecodeLimit) -> Self {
        Self(DecodeErrorKind::Limit(limit))
    }

    /// Return the resource class when a finite limit was exceeded.
    #[must_use]
    pub const fn resource_limit(self) -> Option<DecodeLimit> {
        match self.0 {
            DecodeErrorKind::Limit(limit) => Some(limit),
            _ => None,
        }
    }

    /// Return the selected singular field rejected as a duplicate.
    #[must_use]
    pub const fn duplicate_singular_field(self) -> Option<&'static str> {
        match self.0 {
            DecodeErrorKind::Duplicate(field) => Some(field),
            _ => None,
        }
    }

    /// Return the noncanonical wire detail, when applicable.
    #[must_use]
    pub const fn noncanonical_reason(self) -> Option<&'static str> {
        match self.0 {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.0 {
            DecodeErrorKind::InvalidWire => formatter.write_str("invalid hyperlink wire type"),
            DecodeErrorKind::Truncated => formatter.write_str("truncated hyperlink wire data"),
            DecodeErrorKind::NonCanonical(reason) => {
                write!(
                    formatter,
                    "non-canonical hyperlink wire representation: {reason}"
                )
            },
            DecodeErrorKind::InvalidUtf8(field) => {
                write!(formatter, "invalid UTF-8 in hyperlink field {field}")
            },
            DecodeErrorKind::Duplicate(field) => {
                write!(formatter, "duplicate hyperlink singular field {field}")
            },
            DecodeErrorKind::UnexpectedEndGroup => {
                formatter.write_str("unexpected hyperlink protobuf end-group")
            },
            DecodeErrorKind::MissingEndGroup => {
                formatter.write_str("unterminated hyperlink protobuf group")
            },
            DecodeErrorKind::Limit(DecodeLimit::Bytes { .. }) => {
                formatter.write_str("hyperlink message byte limit exceeded")
            },
            DecodeErrorKind::Limit(DecodeLimit::Fields { .. }) => {
                formatter.write_str("hyperlink field limit exceeded")
            },
            DecodeErrorKind::Limit(DecodeLimit::Work { .. }) => {
                formatter.write_str("hyperlink wire work limit exceeded")
            },
            DecodeErrorKind::Limit(DecodeLimit::Nesting { .. }) => {
                formatter.write_str("hyperlink nesting limit exceeded")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

/// Failure from the bounded canonical hyperlink encoder.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EncodeError {
    /// The generated output would not fit in the platform `usize` or could
    /// not reserve its bounded output allocation.
    Allocation,
}

impl fmt::Display for EncodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("could not allocate hyperlink wire output")
    }
}

impl std::error::Error for EncodeError {}

/// Decode one complete `TSWP.HyperlinkFieldArchive` without generated types.
pub fn decode_hyperlink<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<HyperlinkSnapshot<'source>, DecodeError> {
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::limit(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: options.max_message_bytes,
        }));
    }
    if options.recursion_limit == 0 {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: 1,
            maximum: options.recursion_limit,
        }));
    }
    let mut budget = Budget::new(options);
    let values = scan_message(source, 1, MessageKind::Root, &mut budget)?;
    let MessageValues::Root {
        text_attribute_uuid,
        url_ref,
    } = values
    else {
        return Err(DecodeError::invalid_wire());
    };
    Ok(HyperlinkSnapshot {
        text_attribute_uuid,
        url_ref,
    })
}

/// Encode a canonical hyperlink containing the complete selected schema.
pub fn encode_hyperlink(uuid: &str, url: &str) -> Result<Vec<u8>, EncodeError> {
    let nested_len = field_len(UUID_FIELD, uuid.len())?;
    let total_len = field_len(SUPER_FIELD, nested_len)?
        .checked_add(field_len(URL_FIELD, url.len())?)
        .ok_or(EncodeError::Allocation)?;
    let mut output = Vec::new();
    output
        .try_reserve(total_len)
        .map_err(|_| EncodeError::Allocation)?;
    append_length_delimited(&mut output, SUPER_FIELD, &encode_smart_field(uuid)?);
    append_length_delimited(&mut output, URL_FIELD, url.as_bytes());
    Ok(output)
}

fn encode_smart_field(uuid: &str) -> Result<Vec<u8>, EncodeError> {
    let length = field_len(UUID_FIELD, uuid.len())?;
    let mut output = Vec::new();
    output
        .try_reserve(length)
        .map_err(|_| EncodeError::Allocation)?;
    append_length_delimited(&mut output, UUID_FIELD, uuid.as_bytes());
    Ok(output)
}

fn field_len(field: u32, payload_len: usize) -> Result<usize, EncodeError> {
    let key_len = varint_len(u64::from((field << 3) | 2));
    let payload_len_u64 = u64::try_from(payload_len).map_err(|_| EncodeError::Allocation)?;
    key_len
        .checked_add(varint_len(payload_len_u64))
        .and_then(|length| length.checked_add(payload_len))
        .ok_or(EncodeError::Allocation)
}

fn append_length_delimited(output: &mut Vec<u8>, field: u32, payload: &[u8]) {
    append_varint(output, u64::from((field << 3) | 2));
    append_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

const fn varint_len(value: u64) -> usize {
    if value < 1 << 7 {
        1
    } else if value < 1 << 14 {
        2
    } else if value < 1 << 21 {
        3
    } else if value < 1 << 28 {
        4
    } else if value < 1 << 35 {
        5
    } else if value < 1 << 42 {
        6
    } else if value < 1 << 49 {
        7
    } else if value < 1 << 56 {
        8
    } else if value < 1 << 63 {
        9
    } else {
        10
    }
}

#[derive(Clone, Copy)]
struct Budget {
    fields: usize,
    work_bytes: usize,
    options: DecodeOptions,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            options,
        }
    }

    fn field(&mut self) -> Result<(), DecodeError> {
        self.fields = self.fields.checked_add(1).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: self.options.max_fields,
            })
        })?;
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        Ok(())
    }

    fn work(&mut self, amount: usize) -> Result<(), DecodeError> {
        self.work_bytes = self.work_bytes.checked_add(amount).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: self.options.max_work_bytes,
            })
        })?;
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        Ok(())
    }

    fn depth(&self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        Ok(())
    }
}

#[derive(Clone, Copy)]
enum MessageKind {
    Root,
    Smart,
}

enum MessageValues<'source> {
    Root {
        text_attribute_uuid: Option<&'source str>,
        url_ref: Option<&'source str>,
    },
    Smart {
        text_attribute_uuid: Option<&'source str>,
    },
}

#[derive(Clone, Copy)]
enum Field<'source> {
    Value {
        number: u32,
        wire: WireType,
        bytes: Option<&'source [u8]>,
    },
    StartGroup(u32),
    EndGroup(u32),
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum WireType {
    Varint,
    Fixed64,
    LengthDelimited,
    StartGroup,
    EndGroup,
    Fixed32,
}

fn scan_message<'source>(
    source: &'source [u8],
    depth: u32,
    kind: MessageKind,
    budget: &mut Budget,
) -> Result<MessageValues<'source>, DecodeError> {
    budget.depth(depth)?;
    let mut remaining = source;
    let mut text_attribute_uuid = None;
    let mut url_ref = None;
    let mut seen_super = false;
    let mut seen_url = false;
    let mut seen_uuid = false;
    while !remaining.is_empty() {
        let field = next_field(&mut remaining, budget)?;
        match field {
            Field::EndGroup(_) => {
                return Err(DecodeError::unexpected_end_group());
            },
            Field::StartGroup(number) => {
                skip_group(&mut remaining, depth.saturating_add(1), number, budget)?;
            },
            Field::Value {
                number,
                wire,
                bytes,
            } => match kind {
                MessageKind::Root if number == SUPER_FIELD => {
                    if seen_super {
                        return Err(DecodeError::duplicate("super"));
                    }
                    seen_super = true;
                    if wire != WireType::LengthDelimited {
                        return Err(DecodeError::invalid_wire());
                    }
                    let nested = bytes.ok_or_else(DecodeError::invalid_wire)?;
                    let nested_values =
                        scan_message(nested, depth.saturating_add(1), MessageKind::Smart, budget)?;
                    let MessageValues::Smart {
                        text_attribute_uuid: uuid,
                    } = nested_values
                    else {
                        return Err(DecodeError::invalid_wire());
                    };
                    text_attribute_uuid = uuid;
                },
                MessageKind::Root if number == URL_FIELD => {
                    if seen_url {
                        return Err(DecodeError::duplicate("url_ref"));
                    }
                    seen_url = true;
                    if wire != WireType::LengthDelimited {
                        return Err(DecodeError::invalid_wire());
                    }
                    let value = bytes.ok_or_else(DecodeError::invalid_wire)?;
                    url_ref = Some(
                        str::from_utf8(value).map_err(|_| DecodeError::invalid_utf8("url_ref"))?,
                    );
                },
                MessageKind::Smart if number == UUID_FIELD => {
                    if seen_uuid {
                        return Err(DecodeError::duplicate("text_attribute_uuid_string"));
                    }
                    seen_uuid = true;
                    if wire != WireType::LengthDelimited {
                        return Err(DecodeError::invalid_wire());
                    }
                    let value = bytes.ok_or_else(DecodeError::invalid_wire)?;
                    text_attribute_uuid =
                        Some(str::from_utf8(value).map_err(|_| {
                            DecodeError::invalid_utf8("text_attribute_uuid_string")
                        })?);
                },
                _ => {},
            },
        }
    }
    Ok(match kind {
        MessageKind::Root => MessageValues::Root {
            text_attribute_uuid,
            url_ref,
        },
        MessageKind::Smart => MessageValues::Smart {
            text_attribute_uuid,
        },
    })
}

fn skip_group(
    source: &mut &[u8],
    depth: u32,
    expected_end_group: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.depth(depth)?;
    loop {
        let field = match next_field(source, budget) {
            Ok(field) => field,
            Err(DecodeError(DecodeErrorKind::Truncated)) => {
                return Err(DecodeError::missing_end_group());
            },
            Err(error) => return Err(error),
        };
        match field {
            Field::EndGroup(number) if number == expected_end_group => return Ok(()),
            Field::EndGroup(_) => return Err(DecodeError::unexpected_end_group()),
            Field::StartGroup(number) => {
                skip_group(source, depth.saturating_add(1), number, budget)?;
            },
            Field::Value { .. } => {},
        }
    }
}

fn next_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
) -> Result<Field<'source>, DecodeError> {
    let before = source.len();
    let (raw_key, _) = take_varint(source, "field key")?;
    let number = u32::try_from(raw_key >> 3).map_err(|_| DecodeError::invalid_wire())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid_wire());
    }
    let wire = match raw_key & 7 {
        0 => WireType::Varint,
        1 => WireType::Fixed64,
        2 => WireType::LengthDelimited,
        3 => WireType::StartGroup,
        4 => WireType::EndGroup,
        5 => WireType::Fixed32,
        _ => return Err(DecodeError::invalid_wire()),
    };
    budget.field()?;
    let result = match wire {
        WireType::Varint => {
            let _ = take_varint(source, "varint")?;
            Field::Value {
                number,
                wire,
                bytes: None,
            }
        },
        WireType::Fixed64 => {
            take(source, 8)?;
            Field::Value {
                number,
                wire,
                bytes: None,
            }
        },
        WireType::LengthDelimited => Field::Value {
            number,
            wire,
            bytes: Some(take_length(source)?),
        },
        WireType::StartGroup | WireType::EndGroup => {
            if wire == WireType::StartGroup {
                Field::StartGroup(number)
            } else {
                Field::EndGroup(number)
            }
        },
        WireType::Fixed32 => {
            take(source, 4)?;
            Field::Value {
                number,
                wire,
                bytes: None,
            }
        },
    };
    budget.work(before.saturating_sub(source.len()))?;
    Ok(result)
}

fn take_varint(source: &mut &[u8], label: &'static str) -> Result<(u64, bool), DecodeError> {
    let mut value = 0_u64;
    for index in 0..10 {
        let byte = *source.first().ok_or_else(DecodeError::truncated)?;
        *source = &source[1..];
        if index == 9 && byte > 1 {
            return Err(DecodeError::noncanonical(label));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let canonical = varint_len(value) == index + 1;
            if !canonical {
                return Err(DecodeError::noncanonical(label));
            }
            return Ok((value, canonical));
        }
    }
    Err(DecodeError::noncanonical(label))
}

fn take_length<'source>(source: &mut &'source [u8]) -> Result<&'source [u8], DecodeError> {
    let (length, _) = take_varint(source, "length")?;
    let length = usize::try_from(length).map_err(|_| DecodeError::invalid_wire())?;
    take(source, length)
}

fn take<'source>(source: &mut &'source [u8], length: usize) -> Result<&'source [u8], DecodeError> {
    if source.len() < length {
        return Err(DecodeError::truncated());
    }
    let (selected, remaining) = source.split_at(length);
    *source = remaining;
    Ok(selected)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
    }

    #[test]
    fn canonical_payload_exposes_borrowed_values_and_round_trips_encoding() {
        let source = encode_hyperlink(
            "00112233-4455-6677-8899-aabbccddeeff",
            "https://example.test",
        )
        .expect("canonical hyperlink encoding");
        let snapshot = decode_hyperlink(&source, options(&source)).expect("strict decode");
        assert_eq!(
            snapshot.text_attribute_uuid(),
            Some("00112233-4455-6677-8899-aabbccddeeff")
        );
        assert_eq!(snapshot.url_ref(), Some("https://example.test"));
    }

    #[test]
    fn unknown_fields_are_framed_without_materialization() {
        let mut source = encode_hyperlink("uuid", "url").expect("canonical encoding");
        source.extend_from_slice(&[0x18, 0x01]); // unknown varint field 3
        source.extend_from_slice(&[0x23, 0x28, 0x01, 0x24]); // unknown field-4 group
        let snapshot = decode_hyperlink(&source, options(&source)).expect("unknowns are opaque");
        assert_eq!(snapshot.url_ref(), Some("url"));
    }

    #[test]
    fn malformed_selected_wire_is_rejected() {
        for source in [
            vec![0x08, 0x01],                                     // super has varint wire
            vec![0x0a, 0x03, 0x0a, 0x01],                         // truncated nested message
            vec![0x12, 0x04, b'u', b'r'],                         // truncated URL
            vec![0x0a, 0x02, 0x0a, 0x01, 0x01],                   // invalid nested UTF-8
            vec![0x12, 0x01, 0xff],                               // invalid URL UTF-8
            vec![0x0a, 0x03, 0x0a, 0x01, b'u', 0x0a, 0x01, b'v'], // duplicate UUID
        ] {
            assert!(
                decode_hyperlink(&source, options(&source)).is_err(),
                "{source:?}"
            );
        }
    }

    #[test]
    fn duplicate_and_noncanonical_selected_fields_fail_closed() {
        let duplicate_url = [0x12, 0x01, b'a', 0x12, 0x01, b'b'];
        let error = decode_hyperlink(&duplicate_url, options(&duplicate_url)).unwrap_err();
        assert_eq!(error.duplicate_singular_field(), Some("url_ref"));

        // The canonical key for field 2 is 0x12; the same key with a
        // redundant continuation byte must not be accepted.
        let noncanonical = [0x92, 0x80, 0x00, 0x01, b'a'];
        assert!(decode_hyperlink(&noncanonical, options(&noncanonical)).is_err());
    }

    #[test]
    fn unterminated_groups_and_limits_are_rejected() {
        let grouped = [0x23, 0x28, 0x01];
        assert!(matches!(
            decode_hyperlink(&grouped, options(&grouped)),
            Err(DecodeError(DecodeErrorKind::MissingEndGroup))
        ));
        let source = encode_hyperlink("uuid", "url").expect("canonical encoding");
        let limited = DecodeOptions::new(source.len() - 1, usize::MAX, usize::MAX, 8);
        assert!(matches!(
            decode_hyperlink(&source, limited)
                .unwrap_err()
                .resource_limit(),
            Some(DecodeLimit::Bytes { .. })
        ));
        let limited = DecodeOptions::new(source.len(), 1, usize::MAX, 8);
        assert!(matches!(
            decode_hyperlink(&source, limited)
                .unwrap_err()
                .resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
    }

    #[test]
    fn source_production_has_no_generated_codec_operations() {
        let source = include_str!("hyperlink_codec.rs")
            .split_once("#[cfg(test)]")
            .expect("test module marker")
            .0;
        for marker in ["prost::", "Message::decode", "encode_to_vec", "try_encode"] {
            assert!(!source.contains(marker), "production marker {marker}");
        }
    }
}
