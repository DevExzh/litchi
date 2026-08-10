//! Private-type Buffa projection for the Keynote root document.
//!
//! The generated projection contains only `KN.DocumentArchive.show` and its
//! three scalar `TSP.Reference` fields. The caller must first validate the root
//! document's required opaque base envelope and the uniqueness of the show
//! reference. Unknown bytes remain in the caller-owned source and are never
//! materialized or re-encoded here.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict projection types intentionally stay beside the passes that consume them."
)]

use std::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_document_generated::LitchiIwaProjection as projection;

const DOCUMENT_SHOW_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MAX_RECURSION_LIMIT: u32 = 64;

/// Finite limits already established by the Keynote root wire preflight.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}

/// Generated-free exact `KN.DocumentArchive.show` reference facts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ShowReferenceSnapshot {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

impl ShowReferenceSnapshot {
    /// Required native show identifier.
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }

    /// Optional legacy type discriminator, preserved exactly.
    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }

    /// Optional legacy external marker, preserved exactly.
    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }
}

impl DecodeOptions {
    /// Build an explicit finite profile for one preflighted Keynote root.
    ///
    /// The compatibility constructor derives conservative field and aggregate
    /// scan-work ceilings from the byte ceiling. Callers that already have
    /// independent wire limits can replace those two ceilings with the
    /// `with_max_*` builders.
    #[must_use]
    pub const fn new(max_message_bytes: usize, recursion_limit: u32) -> Self {
        Self {
            max_message_bytes,
            max_fields: max_message_bytes.saturating_mul(4),
            max_work_bytes: max_message_bytes.saturating_mul(8),
            recursion_limit,
        }
    }

    /// Replace the aggregate strict-preflight field ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, max_fields: usize) -> Self {
        self.max_fields = max_fields;
        self
    }

    /// Replace the aggregate strict-plus-Buffa scan-work ceiling in bytes.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, max_work_bytes: usize) -> Self {
        self.max_work_bytes = max_work_bytes;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }

    fn descend(self, budget: &Budget) -> Result<Self, DecodeError> {
        if self.recursion_limit <= 1 {
            return Err(budget.nesting_limit());
        }
        Ok(Self {
            recursion_limit: self.recursion_limit - 1,
            ..self
        })
    }
}

/// Failure from the private Keynote document projection decoder.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

/// A content-free byte or nesting resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    /// Complete input bytes exceeded the configured or Buffa hard ceiling.
    Bytes {
        /// Exact input or configured value that exceeded the ceiling.
        observed: usize,
        /// Exact applied ceiling.
        maximum: usize,
    },
    /// Configured or traversed nesting exceeded its exact ceiling.
    Nesting {
        /// Exact configured value or first rejected depth.
        observed: u32,
        /// Exact applied ceiling.
        maximum: u32,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Resource(WireResourceLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
    Projection,
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Resource(WireResourceLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "Keynote root projection byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(WireResourceLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Keynote root projection nesting limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Keynote root projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Keynote root projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Projection => {
                formatter.write_str("Keynote root strict preflight disagrees with Buffa projection")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "All present and future non-resource Buffa failures retain their exact wire error."
)]
impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self {
            kind: match error {
                buffa::DecodeError::MessageTooLarge
                | buffa::DecodeError::RecursionLimitExceeded => DecodeErrorKind::Projection,
                other => DecodeErrorKind::Wire(other),
            },
        }
    }
}

impl DecodeError {
    /// Required schema field absent from the source, when applicable.
    #[must_use]
    pub const fn missing_required(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Singular field repeated during strict preflight, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        let DecodeErrorKind::DuplicateSingular(field) = self.kind else {
            return None;
        };
        Some(field)
    }

    /// Stable canonicality failure reason, when applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        let DecodeErrorKind::NonCanonical(reason) = self.kind else {
            return None;
        };
        Some(reason)
    }

    /// Exact observed and configured field counts for a field-limit failure.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::FieldLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Exact observed and configured work bytes for a work-limit failure.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::WorkLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Exact byte or nesting resource failure, when applicable.
    #[must_use]
    pub const fn wire_resource_limit(&self) -> Option<WireResourceLimit> {
        let DecodeErrorKind::Resource(limit) = self.kind else {
            return None;
        };
        Some(limit)
    }

    const fn missing_required_field(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::MissingRequired(field),
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

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }
}

/// Decode the show identifier from one already-preflighted Keynote root.
///
/// The singular deferred show reference is always accessed, forcing Buffa to
/// validate its complete wire payload and required identifier before the
/// scalar is returned. The generated view and all unknown root fields remain
/// private and borrowed.
pub fn decode_show_identifier(source: &[u8], options: DecodeOptions) -> Result<u64, DecodeError> {
    Ok(decode_show_reference(source, options)?.identifier())
}

/// Decode the complete generated-free root show reference.
///
/// This forces Buffa's lazy singular reference and retains the exact optional
/// type/external presence alongside the required identifier.
pub fn decode_show_reference(
    source: &[u8],
    options: DecodeOptions,
) -> Result<ShowReferenceSnapshot, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_document(source, options, &mut budget)?;
    budget.charge_work(source.len())?;
    let view: projection::KeynoteDocumentArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    if !view.has_show() {
        return Err(DecodeError::missing_required_field(
            "KN.DocumentArchive.show",
        ));
    }
    budget.charge_work(strict.reference_bytes)?;
    let show = view
        .show
        .get()?
        .ok_or_else(|| DecodeError::missing_required_field("KN.DocumentArchive.show"))?;
    if !show.has_identifier() {
        return Err(DecodeError::missing_required_field(
            "TSP.Reference.identifier",
        ));
    }
    let projected = ShowReferenceSnapshot {
        identifier: show.identifier,
        deprecated_type: show.deprecated_type,
        deprecated_is_external: show.deprecated_is_external,
    };
    if projected != strict.snapshot {
        return Err(DecodeError::projection());
    }
    Ok(strict.snapshot)
}

#[derive(Debug, Clone, Copy)]
struct StrictDocument {
    snapshot: ShowReferenceSnapshot,
    reference_bytes: usize,
}

#[derive(Debug)]
struct Budget {
    fields: usize,
    work_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_nesting: u32,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
            max_nesting: options.recursion_limit,
        }
    }

    fn charge_field(&mut self) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(1);
        if observed > self.max_fields {
            return Err(DecodeError {
                kind: DecodeErrorKind::FieldLimit {
                    observed,
                    maximum: self.max_fields,
                },
            });
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_work(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self.work_bytes.saturating_add(bytes);
        if observed > self.max_work_bytes {
            return Err(DecodeError {
                kind: DecodeErrorKind::WorkLimit {
                    observed,
                    maximum: self.max_work_bytes,
                },
            });
        }
        self.work_bytes = observed;
        Ok(())
    }

    const fn nesting_limit(&self) -> DecodeError {
        DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Nesting {
                observed: self.max_nesting.saturating_add(1),
                maximum: self.max_nesting,
            }),
        }
    }
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard_maximum =
        usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| DecodeError::projection())?;
    if options.max_message_bytes > hard_maximum {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Bytes {
                observed: options.max_message_bytes,
                maximum: hard_maximum,
            }),
        });
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Bytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }),
        });
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION_LIMIT,
            }),
        });
    }
    Ok(())
}

fn preflight_document(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<StrictDocument, DecodeError> {
    budget.charge_work(source.len())?;
    let nested = options.descend(budget)?;
    let mut show = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        if field.number != DOCUMENT_SHOW_FIELD {
            continue;
        }
        if show.is_some() {
            return Err(DecodeError::duplicate_singular("KN.DocumentArchive.show"));
        }
        let payload = field.length_delimited()?;
        show = Some(StrictDocument {
            snapshot: preflight_reference(payload, nested, budget)?,
            reference_bytes: payload.len(),
        });
    }
    show.ok_or_else(|| DecodeError::missing_required_field("KN.DocumentArchive.show"))
}

fn preflight_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ShowReferenceSnapshot, DecodeError> {
    budget.charge_work(source.len())?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
                }
                identifier = Some(field.varint()?);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_type",
                    ));
                }
                deprecated_type = Some(require_canonical_int32(field.varint()?)?);
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if deprecated_is_external.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                deprecated_is_external = Some(require_canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    Ok(ShowReferenceSnapshot {
        identifier: identifier
            .ok_or_else(|| DecodeError::missing_required_field("TSP.Reference.identifier"))?,
        deprecated_type,
        deprecated_is_external,
    })
}

fn require_canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}

fn require_canonical_int32(value: u64) -> Result<i32, DecodeError> {
    const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;
    if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical(
            "int32 scalar is not a sign-extended 32-bit value",
        ));
    }
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        reason = "The strict range check proves canonical int32 sign extension."
    )]
    let decoded = value as i32;
    Ok(decoded)
}

#[derive(Debug, Clone, Copy)]
enum StrictValue<'source> {
    Varint(u64),
    Fixed64,
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32,
}

#[derive(Debug, Clone, Copy)]
struct StrictField<'source> {
    number: u32,
    wire_type: buffa::encoding::WireType,
    value: StrictValue<'source>,
    canonical_key: bool,
    canonical_value: bool,
}

impl<'source> StrictField<'source> {
    fn require_wire_type(self, expected: buffa::encoding::WireType) -> Result<(), DecodeError> {
        if !self.canonical_key {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
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
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("protobuf varint value"));
        }
        let StrictValue::Varint(value) = self.value else {
            return Err(DecodeError::projection());
        };
        Ok(value)
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::LengthDelimited)?;
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("length-delimited size"));
        }
        let StrictValue::LengthDelimited(value) = self.value else {
            return Err(DecodeError::projection());
        };
        Ok(value)
    }
}

#[derive(Debug, Clone, Copy)]
enum ParseItem<'source> {
    Field(StrictField<'source>),
    EndGroup(u32),
}

fn next_strict_field<'source>(
    source: &mut &'source [u8],
    recursion_limit: u32,
    budget: &mut Budget,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_strict_field(source, recursion_limit, budget)? {
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
    budget.charge_field()?;
    let raw_tag =
        u32::try_from(encoded_tag).map_err(|_error| buffa::DecodeError::InvalidFieldNumber)?;
    let field_number = raw_tag >> 3;
    if field_number == 0 || field_number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let raw_wire_type = raw_tag & 7;
    let wire_type = buffa::encoding::WireType::from_u32(raw_wire_type)?;
    let (value, canonical_value) = match wire_type {
        buffa::encoding::WireType::Varint => {
            let (value, canonical) = take_varint(source)?;
            (StrictValue::Varint(value), canonical)
        },
        buffa::encoding::WireType::Fixed64 => {
            let _bytes = take_exact(source, 8)?;
            (StrictValue::Fixed64, true)
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (encoded_length, canonical) = take_varint(source)?;
            let length = usize::try_from(encoded_length)
                .map_err(|_error| buffa::DecodeError::MessageTooLarge)?;
            (
                StrictValue::LengthDelimited(take_exact(source, length)?),
                canonical,
            )
        },
        buffa::encoding::WireType::StartGroup => {
            let child_limit = recursion_limit
                .checked_sub(1)
                .ok_or_else(|| budget.nesting_limit())?;
            skip_strict_group(source, field_number, child_limit, budget)?;
            (StrictValue::Group, true)
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            let _bytes = take_exact(source, 4)?;
            (StrictValue::Fixed32, true)
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw_wire_type).into()),
    };
    Ok(Some(ParseItem::Field(StrictField {
        number: field_number,
        wire_type,
        value,
        canonical_key,
        canonical_value,
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
    reason = "Focused negative tests attach an explicit failure reason to each expected error."
)]
mod tests {
    use prost::Message as _;

    use super::{DecodeOptions, WireResourceLimit, decode_show_identifier, decode_show_reference};

    fn decode(source: &[u8]) -> Result<u64, super::DecodeError> {
        decode_show_identifier(source, DecodeOptions::new(source.len(), 2))
    }

    #[test]
    fn opaque_document_super_is_not_decoded() -> Result<(), Box<dyn std::error::Error>> {
        let source = [0x12, 0x02, 0x08, 0x2a, 0x1a, 0x01, 0xff];
        assert_eq!(decode(&source)?, 42);
        Ok(())
    }

    #[test]
    fn canonical_prost_document_matches_the_projection() -> Result<(), Box<dyn std::error::Error>> {
        let source = crate::kn::DocumentArchive {
            super_: crate::tsa::DocumentArchive::default(),
            show: crate::tsp::Reference {
                identifier: 42,
                deprecated_type: Some(7),
                deprecated_is_external: Some(false),
            },
            tables_custom_format_list: None,
        }
        .encode_to_vec();

        assert_eq!(decode(&source)?, 42);
        Ok(())
    }

    #[test]
    fn nested_reference_is_forced_and_required() {
        let Err(error) = decode(&[0x12, 0x00, 0x1a, 0x00]) else {
            panic!("a show reference without its required identifier must fail");
        };
        assert_eq!(error.missing_required(), Some("TSP.Reference.identifier"));
    }

    #[test]
    fn malformed_nested_reference_is_rejected() {
        let Err(error) = decode(&[0x12, 0x01, 0x08, 0x1a, 0x00]) else {
            panic!("the deferred show reference must be decoded");
        };
        assert!(error.missing_required().is_none());
    }

    #[test]
    fn missing_show_is_rejected() {
        let Err(error) = decode(&[0x1a, 0x00]) else {
            panic!("the required show must be present");
        };
        assert_eq!(error.missing_required(), Some("KN.DocumentArchive.show"));
    }

    #[test]
    fn exact_resource_boundaries_are_typed_and_inclusive() -> Result<(), Box<dyn std::error::Error>>
    {
        const EXACT_FIELDS: usize = 4;
        const EXACT_WORK: usize = 28;

        // One outer field plus all three reference fields, including an
        // explicitly present false external marker.
        let source = [0x12, 0x06, 0x08, 0x2a, 0x10, 0x07, 0x18, 0x00];
        let exact = DecodeOptions::new(source.len(), 2)
            .with_max_fields(EXACT_FIELDS)
            .with_max_work_bytes(EXACT_WORK);
        let snapshot = decode_show_reference(&source, exact)?;
        assert_eq!(snapshot.identifier(), 42);
        assert_eq!(snapshot.deprecated_type(), Some(7));
        assert_eq!(snapshot.deprecated_is_external(), Some(false));

        let bytes = decode_show_reference(
            &source,
            DecodeOptions::new(source.len() - 1, 2)
                .with_max_fields(EXACT_FIELDS)
                .with_max_work_bytes(EXACT_WORK),
        )
        .expect_err("one byte below the exact source length");
        assert_eq!(
            bytes.wire_resource_limit(),
            Some(WireResourceLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1,
            })
        );

        let fields = decode_show_reference(
            &source,
            DecodeOptions::new(source.len(), 2)
                .with_max_fields(EXACT_FIELDS - 1)
                .with_max_work_bytes(EXACT_WORK),
        )
        .expect_err("one field below the exact strict visit count");
        assert_eq!(
            fields.field_limit_values(),
            Some((EXACT_FIELDS, EXACT_FIELDS - 1))
        );

        let work = decode_show_reference(
            &source,
            DecodeOptions::new(source.len(), 2)
                .with_max_fields(EXACT_FIELDS)
                .with_max_work_bytes(EXACT_WORK - 1),
        )
        .expect_err("one byte below strict-plus-Buffa work");
        assert_eq!(work.work_limit_values(), Some((EXACT_WORK, EXACT_WORK - 1)));

        let nesting = decode_show_reference(
            &source,
            DecodeOptions::new(source.len(), 1)
                .with_max_fields(EXACT_FIELDS)
                .with_max_work_bytes(EXACT_WORK),
        )
        .expect_err("one level below Document -> Reference");
        assert_eq!(
            nesting.wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 2,
                maximum: 1,
            })
        );
        Ok(())
    }

    #[test]
    fn invalid_profiles_and_structural_adversaries_never_lose_typed_limits() {
        let source = [0x12, 0x02, 0x08, 0x2a];
        let zero_nesting = decode_show_reference(&source, DecodeOptions::new(source.len(), 0))
            .expect_err("zero is not a valid nesting policy");
        assert_eq!(
            zero_nesting.wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 0,
                maximum: 64,
            })
        );
        let excessive_nesting =
            decode_show_reference(&source, DecodeOptions::new(source.len(), 65))
                .expect_err("Buffa's hard nesting maximum is exact");
        assert_eq!(
            excessive_nesting.wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 65,
                maximum: 64,
            })
        );

        let malformed: [&[u8]; 8] = [
            &[0x80],
            &[0x00],
            &[0x0f],
            &[0x12, 0x02, 0x08],
            &[0x0b],
            &[0x0c],
            &[0x12, 0x01, 0x80],
            &[
                0x08, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x02,
            ],
        ];
        for bytes in malformed {
            assert!(
                decode_show_reference(bytes, DecodeOptions::new(bytes.len().max(1), 8)).is_err()
            );
        }
    }
}
