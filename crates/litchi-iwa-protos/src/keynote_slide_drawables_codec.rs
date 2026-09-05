//! Strict, generated-free facts for Keynote slide drawable discovery.
//!
//! The raw pass is deliberately small and canonical: it validates the slide
//! envelope, every selected reference, and aggregate resource usage before a
//! private Buffa lazy view is created. Buffa then validates the same selected
//! fields lazily, including the repeated reference elements. The returned
//! snapshot owns only scalar identifiers; source bytes and generated views do
//! not escape this boundary.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict parser and its lazy-view cross-check intentionally stay together."
)]

use std::{cell::Cell, fmt};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_slide_drawables_generated::LitchiIwaProjection as projection;

const SLIDE_STYLE_FIELD: u32 = 1;
const SLIDE_TRANSITION_FIELD: u32 = 4;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_IN_DOCUMENT_FIELD: u32 = 19;
const SLIDE_DRAWABLES_Z_ORDER_FIELD: u32 = 42;
const TRANSITION_ATTRIBUTES_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Finite bytes, fields, aggregate-work, and nesting limits for one slide.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build an explicit resource profile for one complete slide payload.
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

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Owned scalar facts required by a read-only Keynote catalog.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideDrawablesSnapshot {
    owned_drawables: Vec<u64>,
    drawables_z_order: Vec<u64>,
}

impl SlideDrawablesSnapshot {
    /// Iterate owned drawable identifiers in native wire order.
    #[must_use]
    pub fn owned_drawables(&self) -> impl ExactSizeIterator<Item = u64> + DoubleEndedIterator + '_ {
        self.owned_drawables.iter().copied()
    }

    /// Iterate drawable identifiers in the native z-order list.
    #[must_use]
    pub fn drawables_z_order(
        &self,
    ) -> impl ExactSizeIterator<Item = u64> + DoubleEndedIterator + '_ {
        self.drawables_z_order.iter().copied()
    }

    /// Move the compact lists into the format-owned package adapter.
    #[must_use]
    pub fn into_parts(self) -> (Vec<u64>, Vec<u64>) {
        (self.owned_drawables, self.drawables_z_order)
    }
}

/// A content-free resource classification for [`DecodeError`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecodeLimit {
    /// Complete source or configured Buffa message-size ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Aggregate strict-plus-lazy field count.
    Fields { observed: usize, maximum: usize },
    /// Aggregate strict-plus-lazy scan work in bytes.
    Work { observed: usize, maximum: usize },
    /// Configured or traversed nesting depth.
    Nesting { observed: u32, maximum: u32 },
}

/// Failure from strict slide preflight or the Buffa lazy-view cross-check.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Resource(DecodeLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    ZeroIdentifier,
    Projection,
    Allocation(&'static str),
}

impl DecodeError {
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

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    const fn allocation(resource: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation(resource),
        }
    }

    const fn resource(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Resource(limit),
        }
    }

    /// Return the bounded-resource classification, when applicable.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match &self.kind {
            DecodeErrorKind::Resource(limit) => Some(*limit),
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Resource(DecodeLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "Keynote slide drawable projection byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "Keynote slide drawable projection field limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "Keynote slide drawable projection work limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Keynote slide drawable projection nesting limit exceeded: observed {observed}, maximum {maximum}"
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
            DecodeErrorKind::ZeroIdentifier => {
                formatter.write_str("Keynote slide drawable reference identifier is zero")
            },
            DecodeErrorKind::Projection => formatter.write_str(
                "Keynote slide drawable strict preflight disagrees with the Buffa projection",
            ),
            DecodeErrorKind::Allocation(resource) => {
                write!(formatter, "allocation failed for {resource}")
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

/// Decode only drawable ownership and z-order from one complete
/// `KN.SlideArchive` payload.
pub fn decode_slide_drawables(
    source: &[u8],
    options: DecodeOptions,
) -> Result<SlideDrawablesSnapshot, DecodeError> {
    validate_input(source, options)?;
    let budget = AggregateBudget::new(options);
    budget.message(source.len())?;
    let strict = preflight_slide(source, options, &budget)?;
    let view: projection::SlideArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    force_projection(&view, &strict)?;
    Ok(SlideDrawablesSnapshot {
        owned_drawables: strict.owned_drawables,
        drawables_z_order: strict.drawables_z_order,
    })
}

#[derive(Debug, PartialEq, Eq)]
struct StrictSlide {
    style: u64,
    in_document: bool,
    owned_drawables: Vec<u64>,
    drawables_z_order: Vec<u64>,
}

fn preflight_slide(
    source: &[u8],
    options: DecodeOptions,
    budget: &AggregateBudget,
) -> Result<StrictSlide, DecodeError> {
    let mut parser = Parser::new(source, options, budget, 1)?;
    let mut style = None;
    let mut transition = false;
    let mut in_document = None;
    let mut owned_drawables = Vec::new();
    let mut drawables_z_order = Vec::new();

    while let Some(field) = parser.field()? {
        match field.number {
            SLIDE_STYLE_FIELD => {
                singular(&mut style, "KN.SlideArchive.style")?;
                let payload = field.bytes()?;
                budget.message(payload.len())?;
                style = Some(preflight_reference(payload, options, budget, 2)?);
            },
            SLIDE_TRANSITION_FIELD => {
                if transition {
                    return Err(DecodeError::duplicate("KN.SlideArchive.transition"));
                }
                transition = true;
                preflight_transition(field.bytes()?, options, budget, 2)?;
            },
            SLIDE_OWNED_DRAWABLES_FIELD => {
                let payload = field.bytes()?;
                budget.message(payload.len())?;
                owned_drawables
                    .try_reserve(1)
                    .map_err(|_error| DecodeError::allocation("owned drawable projection"))?;
                owned_drawables.push(preflight_reference(payload, options, budget, 2)?);
            },
            SLIDE_IN_DOCUMENT_FIELD => {
                singular(&mut in_document, "KN.SlideArchive.inDocument")?;
                in_document = Some(canonical_bool(field.varint()?)?);
            },
            SLIDE_DRAWABLES_Z_ORDER_FIELD => {
                let payload = field.bytes()?;
                budget.message(payload.len())?;
                drawables_z_order
                    .try_reserve(1)
                    .map_err(|_error| DecodeError::allocation("drawable z-order projection"))?;
                drawables_z_order.push(preflight_reference(payload, options, budget, 2)?);
            },
            _ => {},
        }
    }

    if !transition {
        return Err(DecodeError::missing("KN.SlideArchive.transition"));
    }
    Ok(StrictSlide {
        style: style.ok_or_else(|| DecodeError::missing("KN.SlideArchive.style"))?,
        in_document: in_document
            .ok_or_else(|| DecodeError::missing("KN.SlideArchive.inDocument"))?,
        owned_drawables,
        drawables_z_order,
    })
}

fn preflight_transition(
    source: &[u8],
    options: DecodeOptions,
    budget: &AggregateBudget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source.len())?;
    let mut parser = Parser::new(source, options, budget, depth)?;
    let mut attributes = false;
    while let Some(field) = parser.field()? {
        if field.number != TRANSITION_ATTRIBUTES_FIELD {
            continue;
        }
        if attributes {
            return Err(DecodeError::duplicate("KN.TransitionArchive.attributes"));
        }
        attributes = true;
        let payload = field.bytes()?;
        budget.message(payload.len())?;
        let mut attributes_parser = Parser::new(payload, options, budget, depth + 1)?;
        while attributes_parser.field()?.is_some() {}
    }
    if attributes {
        Ok(())
    } else {
        Err(DecodeError::missing("KN.TransitionArchive.attributes"))
    }
}

fn preflight_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &AggregateBudget,
    depth: u32,
) -> Result<u64, DecodeError> {
    let mut parser = Parser::new(source, options, budget, depth)?;
    let mut identifier = None;
    let mut deprecated_type = None::<i32>;
    let mut deprecated_external = None::<bool>;
    while let Some(field) = parser.field()? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                singular(&mut identifier, "TSP.Reference.identifier")?;
                let value = field.varint()?;
                if value == 0 {
                    return Err(DecodeError {
                        kind: DecodeErrorKind::ZeroIdentifier,
                    });
                }
                identifier = Some(value);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                singular(&mut deprecated_type, "TSP.Reference.deprecated_type")?;
                deprecated_type = Some(canonical_int32(field.varint()?)?);
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                singular(
                    &mut deprecated_external,
                    "TSP.Reference.deprecated_is_external",
                )?;
                deprecated_external = Some(canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    identifier.ok_or_else(|| DecodeError::missing("TSP.Reference.identifier"))
}

fn force_projection(
    view: &projection::SlideArchiveLazyView<'_>,
    strict: &StrictSlide,
) -> Result<(), DecodeError> {
    let style = view
        .style
        .get()
        .map_err(DecodeError::from)?
        .ok_or_else(|| DecodeError::missing("KN.SlideArchive.style"))?;
    if !style.has_identifier() || style.identifier != strict.style || style.identifier == 0 {
        return Err(DecodeError::projection());
    }

    let transition = view
        .transition
        .get()
        .map_err(DecodeError::from)?
        .ok_or_else(|| DecodeError::missing("KN.SlideArchive.transition"))?;
    transition
        .attributes
        .get()
        .map_err(DecodeError::from)?
        .ok_or_else(|| DecodeError::missing("KN.TransitionArchive.attributes"))?;

    if !view.has_in_document() || view.in_document != strict.in_document {
        return Err(DecodeError::projection());
    }
    if view.owned_drawables.len() != strict.owned_drawables.len()
        || view.drawables_z_order.len() != strict.drawables_z_order.len()
    {
        return Err(DecodeError::projection());
    }

    for (index, reference) in view.owned_drawables.iter().enumerate() {
        let reference = reference.map_err(DecodeError::from)?;
        let identifier = force_reference(&reference)?;
        if strict.owned_drawables[index] != identifier {
            return Err(DecodeError::projection());
        }
    }
    for (index, reference) in view.drawables_z_order.iter().enumerate() {
        let reference = reference.map_err(DecodeError::from)?;
        let identifier = force_reference(&reference)?;
        if strict.drawables_z_order[index] != identifier {
            return Err(DecodeError::projection());
        }
    }
    Ok(())
}

fn force_reference(view: &projection::ReferenceLazyView<'_>) -> Result<u64, DecodeError> {
    if !view.has_identifier() || view.identifier == 0 {
        return Err(DecodeError::projection());
    }
    Ok(view.identifier)
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard_limit = usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| {
        DecodeError::resource(DecodeLimit::Bytes {
            observed: usize::MAX,
            maximum: usize::MAX,
        })
    })?;
    if options.max_message_bytes > hard_limit {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: options.max_message_bytes,
            maximum: hard_limit,
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

fn singular<T>(value: &mut Option<T>, field: &'static str) -> Result<(), DecodeError> {
    if value.is_some() {
        Err(DecodeError::duplicate(field))
    } else {
        Ok(())
    }
}

fn canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}

fn canonical_int32(value: u64) -> Result<i32, DecodeError> {
    if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical(
            "int32 scalar is not sign-extended",
        ));
    }
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        reason = "The strict range check proves canonical int32 sign extension."
    )]
    let value = value as i32;
    Ok(value)
}

#[derive(Clone, Copy)]
struct Field<'source> {
    number: u32,
    wire_type: u8,
    canonical_key: bool,
    canonical_value: bool,
    value: FieldValue<'source>,
}

#[derive(Clone, Copy)]
enum FieldValue<'source> {
    Varint(u64),
    Bytes(&'source [u8]),
    Other,
}

impl<'source> Field<'source> {
    fn varint(self) -> Result<u64, DecodeError> {
        if self.wire_type != 0 {
            return Err(wire_type_mismatch(self.number, 0, self.wire_type));
        }
        if !self.canonical_key {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("protobuf varint value"));
        }
        match self.value {
            FieldValue::Varint(value) => Ok(value),
            FieldValue::Bytes(_) | FieldValue::Other => Err(DecodeError::projection()),
        }
    }

    fn bytes(self) -> Result<&'source [u8], DecodeError> {
        if self.wire_type != 2 {
            return Err(wire_type_mismatch(self.number, 2, self.wire_type));
        }
        if !self.canonical_key {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("length-delimited size"));
        }
        match self.value {
            FieldValue::Bytes(value) => Ok(value),
            FieldValue::Varint(_) | FieldValue::Other => Err(DecodeError::projection()),
        }
    }
}

fn wire_type_mismatch(field_number: u32, expected: u8, actual: u8) -> DecodeError {
    buffa::DecodeError::WireTypeMismatch {
        field_number,
        expected,
        actual,
    }
    .into()
}

struct Parser<'source, 'budget> {
    remaining: &'source [u8],
    budget: &'budget AggregateBudget,
    depth: u32,
    recursion_limit: u32,
}

impl<'source, 'budget> Parser<'source, 'budget> {
    fn new(
        source: &'source [u8],
        options: DecodeOptions,
        budget: &'budget AggregateBudget,
        depth: u32,
    ) -> Result<Self, DecodeError> {
        budget.depth(depth, options.recursion_limit)?;
        Ok(Self {
            remaining: source,
            budget,
            depth,
            recursion_limit: options.recursion_limit,
        })
    }

    fn field(&mut self) -> Result<Option<Field<'source>>, DecodeError> {
        match parse_field(
            &mut self.remaining,
            self.depth,
            self.recursion_limit,
            self.budget,
        )? {
            Some(ParseItem::Field(field)) => Ok(Some(field)),
            Some(ParseItem::EndGroup(number)) => {
                Err(buffa::DecodeError::InvalidEndGroup(number).into())
            },
            None => Ok(None),
        }
    }
}

enum ParseItem<'source> {
    Field(Field<'source>),
    EndGroup(u32),
}

fn parse_field<'source>(
    source: &mut &'source [u8],
    depth: u32,
    recursion_limit: u32,
    budget: &AggregateBudget,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    let (tag, canonical_key) = take_varint(source)?;
    budget.field()?;
    let field_number =
        u32::try_from(tag >> 3).map_err(|_error| buffa::DecodeError::InvalidFieldNumber)?;
    if field_number == 0 || field_number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let wire_type =
        u8::try_from(tag & 7).map_err(|_error| buffa::DecodeError::InvalidWireType(7))?;
    let value = match wire_type {
        0 => {
            let (value, canonical) = take_varint(source)?;
            return Ok(Some(ParseItem::Field(Field {
                number: field_number,
                wire_type,
                canonical_key,
                canonical_value: canonical,
                value: FieldValue::Varint(value),
            })));
        },
        1 => {
            take_exact(source, 8)?;
            FieldValue::Other
        },
        2 => {
            let (length, canonical) = take_varint(source)?;
            let length =
                usize::try_from(length).map_err(|_error| buffa::DecodeError::MessageTooLarge)?;
            return Ok(Some(ParseItem::Field(Field {
                number: field_number,
                wire_type,
                canonical_key,
                canonical_value: canonical,
                value: FieldValue::Bytes(take_exact(source, length)?),
            })));
        },
        3 => {
            let child_depth = depth.checked_add(1).ok_or_else(DecodeError::projection)?;
            budget.depth(child_depth, recursion_limit)?;
            skip_group(source, field_number, child_depth, recursion_limit, budget)?;
            FieldValue::Other
        },
        4 => return Ok(Some(ParseItem::EndGroup(field_number))),
        5 => {
            take_exact(source, 4)?;
            FieldValue::Other
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(u32::from(wire_type)).into()),
    };
    Ok(Some(ParseItem::Field(Field {
        number: field_number,
        wire_type,
        canonical_key,
        canonical_value: true,
        value,
    })))
}

fn skip_group(
    source: &mut &[u8],
    expected_number: u32,
    depth: u32,
    recursion_limit: u32,
    budget: &AggregateBudget,
) -> Result<(), DecodeError> {
    loop {
        match parse_field(source, depth, recursion_limit, budget)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected_number => return Ok(()),
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
            return Ok((value, canonical_varint_len(value) == consumed));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}

fn canonical_varint_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

struct AggregateBudget {
    fields: Cell<usize>,
    work_bytes: Cell<usize>,
    max_fields: usize,
    max_work_bytes: usize,
}

impl AggregateBudget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: Cell::new(0),
            work_bytes: Cell::new(0),
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
        }
    }

    fn field(&self) -> Result<(), DecodeError> {
        // The strict parser visits each bounded field once and Buffa's lazy
        // view visits that same closure once more. Charge both passes before
        // any deferred access so the aggregate field ceiling is independent
        // of access order.
        let observed = self.fields.get().checked_add(2).ok_or_else(|| {
            DecodeError::resource(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: self.max_fields,
            })
        })?;
        if observed > self.max_fields {
            return Err(DecodeError::resource(DecodeLimit::Fields {
                observed,
                maximum: self.max_fields,
            }));
        }
        self.fields.set(observed);
        Ok(())
    }

    fn message(&self, bytes: usize) -> Result<(), DecodeError> {
        let charge = bytes.checked_mul(2).ok_or_else(|| {
            DecodeError::resource(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: self.max_work_bytes,
            })
        })?;
        let observed = self.work_bytes.get().checked_add(charge).ok_or_else(|| {
            DecodeError::resource(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: self.max_work_bytes,
            })
        })?;
        if observed > self.max_work_bytes {
            return Err(DecodeError::resource(DecodeLimit::Work {
                observed,
                maximum: self.max_work_bytes,
            }));
        }
        self.work_bytes.set(observed);
        Ok(())
    }

    fn depth(&self, depth: u32, maximum: u32) -> Result<(), DecodeError> {
        if depth > maximum {
            return Err(DecodeError::resource(DecodeLimit::Nesting {
                observed: depth,
                maximum,
            }));
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            128,
            source.len().saturating_mul(16).max(1),
            8,
        )
    }

    fn varint(output: &mut Vec<u8>, mut value: u64) {
        while value >= 0x80 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn varint_field(output: &mut Vec<u8>, field: u32, value: u64) {
        varint(output, u64::from(field) << 3);
        varint(output, value);
    }

    fn bytes_field(output: &mut Vec<u8>, field: u32, payload: &[u8]) {
        varint(output, (u64::from(field) << 3) | 2);
        varint(output, payload.len() as u64);
        output.extend_from_slice(payload);
    }

    fn unknown_group(output: &mut Vec<u8>, field: u32, depth: usize) {
        for _ in 0..depth {
            varint(output, (u64::from(field) << 3) | 3);
        }
        for _ in 0..depth {
            varint(output, (u64::from(field) << 3) | 4);
        }
    }

    fn slide(owned: &[u64], z_order: &[u64]) -> Vec<u8> {
        let mut output = Vec::new();
        bytes_field(&mut output, SLIDE_STYLE_FIELD, &[0x08, 0x01]);
        bytes_field(&mut output, SLIDE_TRANSITION_FIELD, &[0x12, 0x00]);
        for &identifier in owned {
            let mut reference = Vec::new();
            varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, identifier);
            bytes_field(&mut output, SLIDE_OWNED_DRAWABLES_FIELD, &reference);
        }
        varint_field(&mut output, SLIDE_IN_DOCUMENT_FIELD, 1);
        for &identifier in z_order {
            let mut reference = Vec::new();
            varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, identifier);
            bytes_field(&mut output, SLIDE_DRAWABLES_Z_ORDER_FIELD, &reference);
        }
        output
    }

    #[test]
    fn canonical_slide_returns_compact_identifier_lists() {
        let source = slide(&[5, 6], &[6, 5]);
        let snapshot = decode_slide_drawables(&source, options(&source)).expect("valid slide");
        assert_eq!(snapshot.owned_drawables().collect::<Vec<_>>(), [5, 6]);
        assert_eq!(snapshot.drawables_z_order().collect::<Vec<_>>(), [6, 5]);
    }

    #[test]
    fn required_envelopes_and_reference_ids_are_strict() {
        let mut missing_transition = slide(&[5], &[5]);
        missing_transition.retain(|byte| *byte != 0x22);
        assert!(decode_slide_drawables(&missing_transition, options(&missing_transition)).is_err());

        let zero = slide(&[0], &[]);
        assert!(decode_slide_drawables(&zero, options(&zero)).is_err());

        let duplicate_style = {
            let mut value = slide(&[], &[]);
            bytes_field(&mut value, SLIDE_STYLE_FIELD, &[0x08, 0x02]);
            value
        };
        assert!(decode_slide_drawables(&duplicate_style, options(&duplicate_style)).is_err());
    }

    #[test]
    fn reference_legacy_fields_remain_unique_and_canonical() {
        let mut duplicate_type = slide(&[], &[]);
        bytes_field(
            &mut duplicate_type,
            SLIDE_OWNED_DRAWABLES_FIELD,
            &[0x08, 0x05, 0x10, 0x01, 0x10, 0x02],
        );
        assert!(decode_slide_drawables(&duplicate_type, options(&duplicate_type)).is_err());

        let mut duplicate_external = slide(&[], &[]);
        bytes_field(
            &mut duplicate_external,
            SLIDE_OWNED_DRAWABLES_FIELD,
            &[0x08, 0x05, 0x18, 0x00, 0x18, 0x01],
        );
        assert!(decode_slide_drawables(&duplicate_external, options(&duplicate_external)).is_err());
    }

    #[test]
    fn unknown_noncanonical_fields_are_skipped_but_selected_fields_are_strict() {
        let mut source = slide(&[5], &[5]);
        // Unknown field 50: overlong key and overlong varint value.
        source.extend_from_slice(&[0x90, 0x83, 0x00, 0x80, 0x00]);
        // Unknown field 51: overlong length prefix for an empty payload.
        source.extend_from_slice(&[0x9a, 0x03, 0x80, 0x00]);
        // Unknown field 52: overlong group key and an overlong nested field.
        source.extend_from_slice(&[0xa3, 0x83, 0x00, 0x88, 0x00, 0x80, 0x00, 0xa4, 0x03]);

        let snapshot = decode_slide_drawables(&source, options(&source))
            .expect("unknown source-owned fields remain opaque");
        assert_eq!(snapshot.owned_drawables().collect::<Vec<_>>(), [5]);
        assert_eq!(snapshot.drawables_z_order().collect::<Vec<_>>(), [5]);

        let mut noncanonical_selected = slide(&[], &[]);
        bytes_field(
            &mut noncanonical_selected,
            SLIDE_OWNED_DRAWABLES_FIELD,
            &[0x08, 0x80, 0x00],
        );
        assert!(
            decode_slide_drawables(&noncanonical_selected, options(&noncanonical_selected))
                .is_err()
        );
    }

    #[test]
    fn aggregate_field_budget_is_applied_before_lazy_view_access() {
        let source = slide(&[5], &[5]);
        let options = DecodeOptions::new(source.len(), 1, source.len() * 16, 8);
        let error = decode_slide_drawables(&source, options).expect_err("field cap");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
    }

    #[test]
    fn message_byte_budget_rejects_a_source_one_byte_over_the_ceiling() {
        let source = slide(&[5], &[5]);
        let options = DecodeOptions::new(source.len() - 1, 128, source.len().saturating_mul(16), 8);
        let error = decode_slide_drawables(&source, options).expect_err("message byte cap");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Bytes { observed, maximum })
                if observed == source.len() && maximum == source.len() - 1
        ));
    }

    #[test]
    fn work_budget_threshold_is_derived_from_an_accepted_decode() {
        let source = slide(&[5, 6], &[6, 5]);
        let decode_with_work =
            |work| decode_slide_drawables(&source, DecodeOptions::new(source.len(), 128, work, 8));

        let mut high = 1usize;
        while decode_with_work(high).is_err() {
            high = high.checked_mul(2).expect("work threshold is finite");
        }
        let mut low = 0usize;
        while low + 1 < high {
            let middle = low + (high - low) / 2;
            if decode_with_work(middle).is_ok() {
                high = middle;
            } else {
                low = middle;
            }
        }

        assert!(decode_with_work(high).is_ok());
        assert!(high > 0);
        let error = decode_with_work(high - 1).expect_err("one byte below work threshold");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Work { observed, maximum })
                if observed > maximum && maximum == high - 1
        ));
    }

    #[test]
    fn nested_unknown_groups_respect_recursion_limits() {
        let mut source = slide(&[5], &[5]);
        unknown_group(&mut source, 50, 4);

        let accepted = DecodeOptions::new(source.len(), 128, source.len() * 16, 64);
        assert!(decode_slide_drawables(&source, accepted).is_ok());

        let bounded = DecodeOptions::new(source.len(), 128, source.len() * 16, 4);
        let error = decode_slide_drawables(&source, bounded).expect_err("nested group depth cap");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting { observed, maximum })
                if observed > maximum && maximum == 4
        ));

        let zero = DecodeOptions::new(source.len(), 128, source.len() * 16, 0);
        assert!(matches!(
            decode_slide_drawables(&source, zero)
                .expect_err("zero recursion limit")
                .resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 0,
                maximum: 64,
            })
        ));
    }

    #[test]
    fn malformed_unknown_groups_are_rejected() {
        let mut truncated = slide(&[5], &[5]);
        varint(&mut truncated, (50_u64 << 3) | 3);
        assert!(decode_slide_drawables(&truncated, options(&truncated)).is_err());

        let mut mismatched = slide(&[5], &[5]);
        varint(&mut mismatched, (50_u64 << 3) | 3);
        varint(&mut mismatched, (51_u64 << 3) | 4);
        assert!(decode_slide_drawables(&mismatched, options(&mismatched)).is_err());
    }
}
