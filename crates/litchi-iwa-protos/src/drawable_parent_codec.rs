//! Strict private Buffa projection for `TSD.DrawableArchive.parent`.
//!
//! Object-index construction visits a drawable payload for every native
//! drawable.  The compatibility path needs only the optional parent edge;
//! decoding the complete generated archive needlessly materializes unrelated
//! geometry, wrapping, accessibility, and annotation state.  This codec
//! performs a bounded canonical wire pass, then cross-checks the selected
//! scalar through a borrowed Buffa lazy view.  The caller-owned payload stays
//! authoritative for unknown fields and future rewrites.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict scanner intentionally precedes the Buffa parity pass."
)]

use std::{fmt, num::NonZeroU64};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_drawable_parent_generated::LitchiIwaProjection as projection;

const DRAWABLE_PARENT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Finite resource policy for one native drawable payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    recursion_maximum: u32,
}

impl DecodeOptions {
    /// Build an explicit finite bytes/fields/work/nesting policy.
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
            recursion_maximum: recursion_limit,
        }
    }

    /// Derive a conservative finite policy from one caller-owned payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(4).max(1),
            bytes.saturating_mul(8).max(1),
            8,
        )
    }

    /// Replace the strict field-visit ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the strict-plus-Buffa work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Replace the protobuf nesting ceiling.
    #[must_use]
    pub const fn with_recursion_limit(mut self, maximum: u32) -> Self {
        self.recursion_limit = maximum;
        self.recursion_maximum = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }

    fn descend(self) -> Result<Self, DecodeError> {
        let recursion_limit = self.recursion_limit.checked_sub(1).ok_or_else(|| {
            DecodeError::recursion_limit(
                self.recursion_maximum.saturating_add(1),
                self.recursion_maximum,
            )
        })?;
        Ok(Self {
            recursion_limit,
            ..self
        })
    }
}

/// Borrow-free selected parent-edge facts.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DrawableParentSnapshot {
    parent: Option<NonZeroU64>,
}

impl DrawableParentSnapshot {
    /// Return the optional non-zero parent object identifier.
    #[must_use]
    pub const fn parent(self) -> Option<NonZeroU64> {
        self.parent
    }

    /// Return the optional parent identifier as its native wire value.
    ///
    /// A zero identifier is not a valid object-index target and is therefore
    /// represented as absent by this semantic projection.
    #[must_use]
    pub const fn parent_identifier(self) -> Option<NonZeroU64> {
        self.parent
    }
}

/// Resource axis rejected by the strict drawable-parent ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Input bytes exceeded the configured ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Encoded fields exceeded the configured ceiling.
    Fields { observed: usize, maximum: usize },
    /// Strict and lazy-view work exceeded the configured ceiling.
    Work { observed: usize, maximum: usize },
    /// Protobuf nesting exceeded the configured ceiling.
    Nesting { observed: u32, maximum: u32 },
}

/// Failure from strict drawable-parent decoding or its Buffa parity check.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Limit(DecodeLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    Projection,
}

impl DecodeError {
    const fn limit(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Limit(limit),
        }
    }

    const fn missing_required(field: &'static str) -> Self {
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

    const fn recursion_limit(observed: u32, maximum: u32) -> Self {
        Self::limit(DecodeLimit::Nesting { observed, maximum })
    }

    /// Return the rejected resource, if this is a finite-limit failure.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Limit(limit) => Some(limit),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Return the missing required field, if applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Limit(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Return the duplicated singular field, if applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Limit(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Return the stable non-canonical wire reason, if applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Limit(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::Projection => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Limit(DecodeLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "TSD drawable-parent projection byte limit exceeded: {observed} > {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "TSD drawable-parent projection field limit exceeded: {observed} > {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "TSD drawable-parent projection work limit exceeded: {observed} > {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "TSD drawable-parent projection nesting limit exceeded: {observed} > {maximum}"
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
            DecodeErrorKind::Projection => formatter.write_str(
                "TSD drawable-parent strict preflight disagrees with the Buffa projection",
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

/// Decode only `TSD.DrawableArchive.parent` through a bounded borrowed view.
///
/// Every source field is still structurally scanned, so unknown fields and
/// groups cannot bypass framing, field, work, or nesting limits. The lazy
/// projection retains no source-owned allocation; only the optional non-zero
/// identifier is returned.
pub fn decode_parent(
    source: &[u8],
    options: DecodeOptions,
) -> Result<DrawableParentSnapshot, DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight(source, options, &mut budget)?;
    let view: projection::DrawableArchiveLazyView<'_> = options.buffa().decode_lazy_view(source)?;
    let parent_view = view.parent.get()?;
    let projected = parent_view.as_ref().map(project_reference).transpose()?;
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(DrawableParentSnapshot {
        parent: strict.and_then(|reference| NonZeroU64::new(reference.identifier)),
    })
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
    if options.max_message_bytes > hard_bytes {
        return Err(DecodeError::limit(DecodeLimit::Bytes {
            observed: options.max_message_bytes,
            maximum: hard_bytes,
        }));
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::limit(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: options.max_message_bytes,
        }));
    }
    if options.recursion_limit == 0 {
        return Err(DecodeError::recursion_limit(1, 0));
    }
    if options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::recursion_limit(
            options.recursion_limit,
            MAX_RECURSION_LIMIT,
        ));
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
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed,
                maximum: self.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_work(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self.work_bytes.saturating_add(bytes.saturating_mul(2));
        if observed > self.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::Work {
                observed,
                maximum: self.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct RawReference {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

fn preflight(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Option<RawReference>, DecodeError> {
    budget.charge_work(source.len())?;
    let mut parent = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(
        &mut remaining,
        options.recursion_limit,
        options.recursion_maximum,
        budget,
    )? {
        if field.number != DRAWABLE_PARENT_FIELD {
            continue;
        }
        if parent.is_some() {
            return Err(DecodeError::duplicate_singular(
                "TSD.DrawableArchive.parent",
            ));
        }
        let nested_options = options.descend()?;
        let payload = field.length_delimited()?;
        parent = Some(preflight_reference(payload, nested_options, budget)?);
    }
    Ok(parent)
}

fn preflight_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<RawReference, DecodeError> {
    budget.charge_work(source.len())?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(
        &mut remaining,
        options.recursion_limit,
        options.recursion_maximum,
        budget,
    )? {
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
    Ok(RawReference {
        identifier: identifier
            .ok_or_else(|| DecodeError::missing_required("TSP.Reference.identifier"))?,
        deprecated_type,
        deprecated_is_external,
    })
}

fn project_reference(
    view: &projection::ReferenceLazyView<'_>,
) -> Result<RawReference, DecodeError> {
    if !view.has_identifier() {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    Ok(RawReference {
        identifier: view.identifier,
        deprecated_type: view.deprecated_type,
        deprecated_is_external: view.deprecated_is_external,
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
    if value <= u64::from(u32::MAX / 2) {
        return i32::try_from(value).map_err(|_conversion| DecodeError::projection());
    }
    if value >= MIN_SIGN_EXTENDED_INT32 {
        let truncated = u32::try_from(value & u64::from(u32::MAX))
            .map_err(|_conversion| DecodeError::projection())?;
        return Ok(i32::from_ne_bytes(truncated.to_ne_bytes()));
    }
    Err(DecodeError::noncanonical(
        "int32 scalar is not sign-extended",
    ))
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
    canonical_key: bool,
    canonical_value: bool,
}

impl<'source> StrictField<'source> {
    fn require_canonical_key(self) -> Result<(), DecodeError> {
        if !self.canonical_key {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
        Ok(())
    }

    fn varint(self) -> Result<u64, DecodeError> {
        self.require_canonical_key()?;
        if self.wire_type != buffa::encoding::WireType::Varint {
            return Err(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected: buffa::encoding::WireType::Varint as u8,
                actual: self.wire_type as u8,
            }
            .into());
        }
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("protobuf varint value"));
        }
        match self.value {
            StrictValue::Varint(value) => Ok(value),
            StrictValue::Fixed64
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group
            | StrictValue::Fixed32 => Err(DecodeError::projection()),
        }
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        self.require_canonical_key()?;
        if self.wire_type != buffa::encoding::WireType::LengthDelimited {
            return Err(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected: buffa::encoding::WireType::LengthDelimited as u8,
                actual: self.wire_type as u8,
            }
            .into());
        }
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("length-delimited size"));
        }
        match self.value {
            StrictValue::LengthDelimited(value) => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64
            | StrictValue::Group
            | StrictValue::Fixed32 => Err(DecodeError::projection()),
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum ParseItem<'source> {
    Field(StrictField<'source>),
    EndGroup(u32),
}

fn next_strict_field<'source>(
    source: &mut &'source [u8],
    recursion_limit: u32,
    recursion_maximum: u32,
    budget: &mut Budget,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_strict_field(source, recursion_limit, recursion_maximum, budget)? {
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
    recursion_maximum: u32,
    budget: &mut Budget,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    let (encoded_tag, canonical_key) = take_varint(source)?;
    budget.charge_field()?;
    let raw_tag =
        u32::try_from(encoded_tag).map_err(|_conversion| buffa::DecodeError::InvalidFieldNumber)?;
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
            take_exact(source, 8)?;
            (StrictValue::Fixed64, true)
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (encoded_length, canonical) = take_varint(source)?;
            let length = usize::try_from(encoded_length)
                .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
            let payload = take_exact(source, length)?;
            (StrictValue::LengthDelimited(payload), canonical)
        },
        buffa::encoding::WireType::StartGroup => {
            let child_limit = recursion_limit.checked_sub(1).ok_or_else(|| {
                DecodeError::recursion_limit(recursion_maximum.saturating_add(1), recursion_maximum)
            })?;
            skip_strict_group(source, field_number, child_limit, recursion_maximum, budget)?;
            (StrictValue::Group, true)
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            take_exact(source, 4)?;
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
    recursion_maximum: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    loop {
        match parse_strict_field(source, recursion_limit, recursion_maximum, budget)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected_field_number => return Ok(()),
            Some(ParseItem::EndGroup(number)) => {
                return Err(buffa::DecodeError::InvalidEndGroup(number).into());
            },
            None => return Err(buffa::DecodeError::UnexpectedEof.into()),
        }
    }
}

fn take_exact<'source>(
    source: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], DecodeError> {
    if source.len() < length {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    let (value, rest) = source.split_at(length);
    *source = rest;
    Ok(value)
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::tsd;
    use prost::Message as _;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
            .with_max_fields(usize::MAX)
            .with_max_work_bytes(usize::MAX)
    }

    fn varint_field(field: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 0);
        push_varint(&mut output, value);
        output
    }

    fn fixed32_field(field: u32, value: u32) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 5);
        output.extend_from_slice(&value.to_le_bytes());
        output
    }

    fn fixed64_field(field: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 1);
        output.extend_from_slice(&value.to_le_bytes());
        output
    }

    fn start_group(field: u32) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 3);
        output
    }

    fn end_group(field: u32) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 4);
        output
    }

    fn length_field(field: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 2);
        push_varint(
            &mut output,
            u64::try_from(payload.len()).expect("fixture length fits u64"),
        );
        output.extend_from_slice(payload);
        output
    }

    fn push_varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = u8::try_from(value & 0x7f).expect("varint chunk fits u8");
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

    fn source_with_parent_values(
        identifier: u64,
        deprecated_type: Option<i32>,
        deprecated_is_external: Option<bool>,
    ) -> Vec<u8> {
        let reference = tsd::DrawableArchive {
            parent: Some(crate::tsp::Reference {
                identifier,
                deprecated_type,
                deprecated_is_external,
            }),
            ..Default::default()
        };
        reference.encode_to_vec()
    }

    fn source_with_parent(identifier: u64) -> Vec<u8> {
        source_with_parent_values(identifier, Some(-7), Some(false))
    }

    #[derive(Debug, PartialEq, Eq)]
    struct ReferenceFacts {
        identifier_present: bool,
        identifier: u64,
        deprecated_type: Option<i32>,
        deprecated_is_external: Option<bool>,
    }

    fn prost_reference_facts(reference: &crate::tsp::Reference) -> ReferenceFacts {
        ReferenceFacts {
            identifier_present: true,
            identifier: reference.identifier,
            deprecated_type: reference.deprecated_type,
            deprecated_is_external: reference.deprecated_is_external,
        }
    }

    fn buffa_reference_facts(reference: projection::ReferenceLazyView<'_>) -> ReferenceFacts {
        ReferenceFacts {
            identifier_present: reference.has_identifier(),
            identifier: reference.identifier,
            deprecated_type: reference.deprecated_type,
            deprecated_is_external: reference.deprecated_is_external,
        }
    }

    fn assert_prost_parent_parity(source: &[u8]) -> Option<ReferenceFacts> {
        let native = tsd::DrawableArchive::decode(source).expect("native decode");
        let view: projection::DrawableArchiveLazyView<'_> = DecodeOptions::for_source(source)
            .with_max_fields(usize::MAX)
            .with_max_work_bytes(usize::MAX)
            .buffa()
            .decode_lazy_view(source)
            .expect("Buffa view");
        let projected = view
            .parent
            .get()
            .expect("parent view")
            .map(buffa_reference_facts);
        let expected = native.parent.as_ref().map(prost_reference_facts);
        assert_eq!(projected, expected);
        expected
    }

    #[test]
    fn parent_projection_matches_every_prost_reference_field() {
        for (identifier, deprecated_type, deprecated_is_external) in [
            (41, Some(-7), Some(false)),
            (0, None, None),
            (7, Some(0), Some(true)),
            (u64::MAX, Some(i32::MIN), Some(true)),
        ] {
            let source =
                source_with_parent_values(identifier, deprecated_type, deprecated_is_external);
            let expected = assert_prost_parent_parity(&source).expect("parent is present");
            assert_eq!(expected.identifier, identifier);
            assert_eq!(expected.deprecated_type, deprecated_type);
            assert_eq!(expected.deprecated_is_external, deprecated_is_external);

            let snapshot = decode_parent(&source, options(&source)).expect("projection");
            assert_eq!(
                snapshot.parent().map(NonZeroU64::get),
                NonZeroU64::new(identifier).map(NonZeroU64::get)
            );
        }
    }

    #[test]
    fn absent_parent_matches_prost_and_returns_absent_snapshot() {
        let source = tsd::DrawableArchive::default().encode_to_vec();
        assert_eq!(assert_prost_parent_parity(&source), None);
        assert_eq!(
            decode_parent(&source, options(&source)).expect("absent parent"),
            DrawableParentSnapshot::default()
        );
    }

    #[test]
    fn unknown_wire_kinds_and_balanced_groups_are_scanned_but_not_retained() {
        let mut source = source_with_parent(9);
        source.extend(varint_field(99, 0xfeed));
        source.extend(fixed32_field(100, 0xfeed_face));
        source.extend(fixed64_field(101, 0xfeed_face_cafe_beef));
        source.extend(length_field(102, b"opaque"));
        source.extend(start_group(103));
        source.extend(varint_field(104, 7));
        source.extend(fixed32_field(105, 0xdecafbad));
        source.extend(start_group(106));
        source.extend(length_field(107, b"nested"));
        source.extend(end_group(106));
        source.extend(end_group(103));

        let snapshot = decode_parent(&source, options(&source)).expect("unknown is opaque");
        assert_eq!(snapshot.parent().map(NonZeroU64::get), Some(9));
    }

    #[test]
    fn malformed_unknown_groups_are_rejected() {
        let cases = [
            (
                [start_group(99), end_group(100)].concat(),
                buffa::DecodeError::InvalidEndGroup(100),
            ),
            (
                [start_group(99), varint_field(100, 7)].concat(),
                buffa::DecodeError::UnexpectedEof,
            ),
            (end_group(99), buffa::DecodeError::InvalidEndGroup(99)),
        ];

        for (source, expected) in cases {
            let error = decode_parent(&source, options(&source)).expect_err("malformed group");
            assert_eq!(error.kind, DecodeErrorKind::Wire(expected));
        }
    }

    #[test]
    fn duplicate_parent_is_rejected_without_publishing() {
        let mut source = source_with_parent(9);
        source.extend(length_field(2, &varint_field(1, 10)));
        let error = decode_parent(&source, options(&source)).expect_err("duplicate parent");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSD.DrawableArchive.parent")
        );
    }

    #[test]
    fn malformed_nested_reference_is_rejected() {
        let source = length_field(2, &varint_field(2, 1));
        let error = decode_parent(&source, options(&source)).expect_err("missing identifier");
        assert_eq!(
            error.missing_required_field(),
            Some("TSP.Reference.identifier")
        );
    }

    #[test]
    fn noncanonical_known_wire_is_rejected() {
        // Re-encode the nested identifier with an overlong varint.
        let source = length_field(2, &[0x08, 0xc9, 0x00]);
        let error = decode_parent(&source, options(&source)).expect_err("overlong value");
        assert_eq!(error.noncanonical_reason(), Some("protobuf varint value"));
    }

    #[test]
    fn finite_field_budget_is_enforced_before_lazy_decode() {
        let source = source_with_parent(3);
        let error = decode_parent(
            &source,
            DecodeOptions::for_source(&source).with_max_fields(1),
        )
        .expect_err("field budget");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
    }

    #[test]
    fn exact_and_one_under_message_byte_limits_are_enforced() {
        let source = source_with_parent(3);
        let exact = DecodeOptions::new(source.len(), usize::MAX, usize::MAX, 1);
        assert!(decode_parent(&source, exact).is_ok());

        let one_under = DecodeOptions::new(source.len() - 1, usize::MAX, usize::MAX, 1);
        let error = decode_parent(&source, one_under).expect_err("one byte under");
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1,
            })
        );
    }

    #[test]
    fn exact_and_one_under_field_limits_are_enforced() {
        let source = source_with_parent(3);
        let exact = DecodeOptions::new(source.len(), 4, usize::MAX, 1);
        assert!(decode_parent(&source, exact).is_ok());

        let one_under = DecodeOptions::new(source.len(), 3, usize::MAX, 1);
        let error = decode_parent(&source, one_under).expect_err("one field under");
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Fields {
                observed: 4,
                maximum: 3,
            })
        );
    }

    #[test]
    fn exact_and_one_under_work_limits_are_enforced() {
        let identifier = 3;
        let reference = crate::tsp::Reference {
            identifier,
            deprecated_type: Some(-7),
            deprecated_is_external: Some(false),
        };
        let nested = reference.encode_to_vec();
        let source = source_with_parent(identifier);
        let exact_work = source.len().saturating_add(nested.len()).saturating_mul(2);
        let exact = DecodeOptions::new(source.len(), usize::MAX, exact_work, 1);
        assert!(decode_parent(&source, exact).is_ok());

        let one_under = DecodeOptions::new(source.len(), usize::MAX, exact_work - 1, 1);
        let error = decode_parent(&source, one_under).expect_err("one work byte under");
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Work {
                observed: exact_work,
                maximum: exact_work - 1,
            })
        );
    }

    #[test]
    fn exact_and_one_under_recursion_limits_are_enforced() {
        let source = source_with_parent(3);
        let exact = DecodeOptions::new(source.len(), usize::MAX, usize::MAX, 1);
        assert!(decode_parent(&source, exact).is_ok());

        let one_under = DecodeOptions::new(source.len(), usize::MAX, usize::MAX, 0);
        let error = decode_parent(&source, one_under).expect_err("one nesting level under");
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 1,
                maximum: 0,
            })
        );
    }

    #[test]
    fn nesting_limit_reports_the_caller_configured_ceiling() {
        // Field 1 starts a group containing a second group. A recursion
        // ceiling of one allows the outer group but rejects the attempted
        // second level as observed depth two.
        let source = [0x0b, 0x13, 0x14, 0x0c];
        let error = decode_parent(
            &source,
            DecodeOptions::for_source(&source).with_recursion_limit(1),
        )
        .expect_err("nested group must exceed the configured ceiling");
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 2,
                maximum: 1,
            })
        );
    }
}
