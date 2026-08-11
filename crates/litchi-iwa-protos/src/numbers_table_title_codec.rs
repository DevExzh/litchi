//! Strict generated-free routing for Numbers table-title settings.
//!
//! One handwritten pass validates every protobuf record before Buffa observes
//! the payload. Buffa then supplies a borrowed lazy-view cross-check for the
//! three scalar title fields and for each selected `TSP.Reference`. Generated
//! values never escape this module and caller-owned bytes remain authoritative
//! for preservation and rewriting.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict routing helpers stay beside the generated-free snapshots they build."
)]

use core::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_numbers_sheet_order_generated::LitchiIwaProjection as reference_projection;
use crate::buffa_numbers_table_title_generated::LitchiIwaProjection as title_projection;

const TABLE_NAME_ENABLED_FIELD: u32 = 22;
const TABLE_NAME_STYLE_FIELD: u32 = 30;
const TABLE_NAME_HEIGHT_FIELD: u32 = 33;
const TABLE_NAME_SHAPE_STYLE_FIELD: u32 = 36;
const TABLE_NAME_BORDER_ENABLED_FIELD: u32 = 37;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MAX_RECURSION: u32 = 64;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const MIN_SIGN_EXTENDED_I32: u64 = 0xffff_ffff_8000_0000;

/// Finite aggregate policy for one table-model title projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_references: usize,
}

impl DecodeOptions {
    /// Construct an explicit bytes/fields/work/nesting/reference policy.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_references: usize,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_references,
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

/// Successful exact consumption for transaction-budget aggregation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    reference_bytes: usize,
}

impl DecodeReport {
    /// Encoded field records inspected by the strict root/reference traversal.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Bytes inspected by strict and Buffa root/reference passes.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Greatest protobuf message or unknown-group depth reached.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Present table-title style reference occurrences.
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    /// Aggregate bytes inside selected reference envelopes.
    #[must_use]
    pub const fn reference_bytes(self) -> usize {
        self.reference_bytes
    }
}

/// Typed finite resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// The owner payload or configured Buffa message ceiling is too large.
    Bytes { observed: usize, maximum: usize },
    /// Selected style-reference occurrences exceed their finite ceiling.
    References { observed: usize, maximum: usize },
    /// Aggregate encoded fields exceed their finite ceiling.
    Fields { observed: usize, maximum: usize },
    /// Strict plus Buffa traversal work exceeds its finite ceiling.
    Work { observed: usize, maximum: usize },
    /// Configured or traversed nesting exceeds its finite ceiling.
    Nesting { observed: u32, maximum: u32 },
}

/// Strict table-title decode failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    limit: Option<DecodeLimit>,
}

impl DecodeError {
    /// Return the exact resource observation for a limit failure.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        self.limit
    }

    const fn invalid() -> Self {
        Self { limit: None }
    }

    const fn limited(limit: DecodeLimit) -> Self {
        Self { limit: Some(limit) }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid Numbers table-title payload")
    }
}

impl std::error::Error for DecodeError {}

/// Generated-free exact scalar projection of one native `TSP.Reference`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReferenceSnapshot {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

impl ReferenceSnapshot {
    /// Required, non-zero native object identifier.
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }

    /// Optional legacy object-type hint with source presence retained.
    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }

    /// Optional legacy external marker with source presence retained.
    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }
}

/// Presence-preserving title settings and rendering prerequisites.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableTitleSettingsSnapshot {
    table_name_enabled: Option<bool>,
    table_name_style: Option<ReferenceSnapshot>,
    table_name_height_bits: Option<u64>,
    table_name_shape_style: Option<ReferenceSnapshot>,
    table_name_border_enabled: Option<bool>,
}

impl TableTitleSettingsSnapshot {
    /// Optional native title-visibility flag from field 22.
    #[must_use]
    pub const fn table_name_enabled(self) -> Option<bool> {
        self.table_name_enabled
    }

    /// Optional native title paragraph-style reference from field 30.
    #[must_use]
    pub const fn table_name_style(self) -> Option<ReferenceSnapshot> {
        self.table_name_style
    }

    /// Exact IEEE-754 bits of the optional title height from field 33.
    #[must_use]
    pub const fn table_name_height_bits(self) -> Option<u64> {
        self.table_name_height_bits
    }

    /// Optional native title shape-style reference from field 36.
    #[must_use]
    pub const fn table_name_shape_style(self) -> Option<ReferenceSnapshot> {
        self.table_name_shape_style
    }

    /// Optional native title-outline flag from field 37.
    #[must_use]
    pub const fn table_name_border_enabled(self) -> Option<bool> {
        self.table_name_border_enabled
    }
}

/// Decode title settings and prerequisites without exposing generated values.
pub fn decode_table_title_settings(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableTitleSettingsSnapshot, DecodeError> {
    Ok(decode_table_title_settings_with_report(source, options)?.0)
}

/// Decode title settings and return exact aggregate resource consumption.
pub fn decode_table_title_settings_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableTitleSettingsSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    budget.charge_work(source.len())?;
    let strict = strict_root(source, &mut budget)?;

    budget.charge_work(source.len())?;
    let view: title_projection::TableTitleSettingsArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.table_name_enabled != strict.snapshot.table_name_enabled
        || view.table_name_height_bits != strict.snapshot.table_name_height_bits
        || view.table_name_border_enabled != strict.snapshot.table_name_border_enabled
    {
        return Err(DecodeError::invalid());
    }

    if let Some(reference) = strict.table_name_style {
        crosscheck_reference(reference, strict.snapshot.table_name_style, &mut budget)?;
    }
    if let Some(reference) = strict.table_name_shape_style {
        crosscheck_reference(
            reference,
            strict.snapshot.table_name_shape_style,
            &mut budget,
        )?;
    }
    Ok((strict.snapshot, budget.report()))
}

#[derive(Debug, Clone, Copy)]
struct StrictRoot<'source> {
    snapshot: TableTitleSettingsSnapshot,
    table_name_style: Option<&'source [u8]>,
    table_name_shape_style: Option<&'source [u8]>,
}

fn strict_root<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<StrictRoot<'source>, DecodeError> {
    let mut snapshot = TableTitleSettingsSnapshot {
        table_name_enabled: None,
        table_name_style: None,
        table_name_height_bits: None,
        table_name_shape_style: None,
        table_name_border_enabled: None,
    };
    let mut table_name_style = None;
    let mut table_name_shape_style = None;
    let mut remaining = source;
    while let Some(field) = next_root_field(&mut remaining, budget, 1)? {
        match field.number {
            TABLE_NAME_ENABLED_FIELD => {
                if snapshot.table_name_enabled.is_some() {
                    return Err(DecodeError::invalid());
                }
                snapshot.table_name_enabled = Some(canonical_bool(field.varint()?)?);
            },
            TABLE_NAME_STYLE_FIELD => {
                if table_name_style.is_some() {
                    return Err(DecodeError::invalid());
                }
                let payload = field.length_delimited()?;
                budget.charge_reference(payload.len())?;
                budget.charge_work(payload.len())?;
                let reference = strict_reference(payload, budget, 2)?;
                table_name_style = Some(payload);
                snapshot.table_name_style = Some(reference);
            },
            TABLE_NAME_HEIGHT_FIELD => {
                if snapshot.table_name_height_bits.is_some() {
                    return Err(DecodeError::invalid());
                }
                snapshot.table_name_height_bits = Some(field.fixed64()?);
            },
            TABLE_NAME_SHAPE_STYLE_FIELD => {
                if table_name_shape_style.is_some() {
                    return Err(DecodeError::invalid());
                }
                let payload = field.length_delimited()?;
                budget.charge_reference(payload.len())?;
                budget.charge_work(payload.len())?;
                let reference = strict_reference(payload, budget, 2)?;
                table_name_shape_style = Some(payload);
                snapshot.table_name_shape_style = Some(reference);
            },
            TABLE_NAME_BORDER_ENABLED_FIELD => {
                if snapshot.table_name_border_enabled.is_some() {
                    return Err(DecodeError::invalid());
                }
                snapshot.table_name_border_enabled = Some(canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    Ok(StrictRoot {
        snapshot,
        table_name_style,
        table_name_shape_style,
    })
}

fn strict_reference(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<ReferenceSnapshot, DecodeError> {
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    let mut remaining = source;
    while let Some(field) = next_root_field(&mut remaining, budget, depth)? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::invalid());
                }
                let value = field.varint()?;
                if value == 0 {
                    return Err(DecodeError::invalid());
                }
                identifier = Some(value);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type.is_some() {
                    return Err(DecodeError::invalid());
                }
                deprecated_type = Some(canonical_int32(field.varint()?)?);
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if deprecated_is_external.is_some() {
                    return Err(DecodeError::invalid());
                }
                deprecated_is_external = Some(canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    Ok(ReferenceSnapshot {
        identifier: identifier.ok_or_else(DecodeError::invalid)?,
        deprecated_type,
        deprecated_is_external,
    })
}

fn crosscheck_reference(
    source: &[u8],
    strict: Option<ReferenceSnapshot>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let strict = strict.ok_or_else(DecodeError::invalid)?;
    budget.charge_work(source.len())?;
    let view: reference_projection::NumbersSheetReferenceArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if !view.has_identifier()
        || view.identifier != strict.identifier
        || view.deprecated_type != strict.deprecated_type
        || view.deprecated_is_external != strict.deprecated_is_external
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct StrictField<'source> {
    number: u32,
    wire_type: u8,
    value: StrictValue<'source>,
}

impl<'source> StrictField<'source> {
    fn varint(self) -> Result<u64, DecodeError> {
        match self.value {
            StrictValue::Varint(value) if self.wire_type == 0 => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64(_)
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group
            | StrictValue::Fixed32 => Err(DecodeError::invalid()),
        }
    }

    fn fixed64(self) -> Result<u64, DecodeError> {
        match self.value {
            StrictValue::Fixed64(value) if self.wire_type == 1 => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64(_)
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group
            | StrictValue::Fixed32 => Err(DecodeError::invalid()),
        }
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        match self.value {
            StrictValue::LengthDelimited(value) if self.wire_type == 2 => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64(_)
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group
            | StrictValue::Fixed32 => Err(DecodeError::invalid()),
        }
    }
}

#[derive(Clone, Copy)]
enum StrictValue<'source> {
    Varint(u64),
    Fixed64(u64),
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32,
}

enum ParseItem<'source> {
    Field(StrictField<'source>),
    EndGroup(u32),
}

fn next_root_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_field(source, budget, depth)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(_)) => Err(DecodeError::invalid()),
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
    budget.charge_field()?;
    let tag = take_varint(source)?;
    let number = u32::try_from(tag >> 3).map_err(|_conversion| DecodeError::invalid())?;
    let wire_type = u8::try_from(tag & 7).map_err(|_conversion| DecodeError::invalid())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid());
    }
    let value = match wire_type {
        0 => StrictValue::Varint(take_varint(source)?),
        1 => StrictValue::Fixed64(u64::from_le_bytes(
            take(source, 8)?
                .try_into()
                .map_err(|_length| DecodeError::invalid())?,
        )),
        2 => {
            let length = usize::try_from(take_varint(source)?)
                .map_err(|_conversion| DecodeError::invalid())?;
            StrictValue::LengthDelimited(take(source, length)?)
        },
        3 => {
            let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
            skip_group(source, number, budget, child_depth)?;
            StrictValue::Group
        },
        4 => return Ok(Some(ParseItem::EndGroup(number))),
        5 => {
            take(source, 4)?;
            StrictValue::Fixed32
        },
        _ => return Err(DecodeError::invalid()),
    };
    Ok(Some(ParseItem::Field(StrictField {
        number,
        wire_type,
        value,
    })))
}

fn skip_group(
    source: &mut &[u8],
    expected_number: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    loop {
        match parse_field(source, budget, depth)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected_number => return Ok(()),
            Some(ParseItem::EndGroup(_)) | None => return Err(DecodeError::invalid()),
        }
    }
}

fn take<'source>(source: &mut &'source [u8], amount: usize) -> Result<&'source [u8], DecodeError> {
    if source.len() < amount {
        return Err(DecodeError::invalid());
    }
    let (selected, remaining) = source.split_at(amount);
    *source = remaining;
    Ok(selected)
}

fn take_varint(source: &mut &[u8]) -> Result<u64, DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original.get(index).ok_or_else(DecodeError::invalid)?;
        if index == 9 && byte > 1 {
            return Err(DecodeError::invalid());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            if encoded_varint_len(value) != consumed {
                return Err(DecodeError::invalid());
            }
            *source = &original[consumed..];
            return Ok(value);
        }
    }
    Err(DecodeError::invalid())
}

const fn encoded_varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

fn canonical_int32(value: u64) -> Result<i32, DecodeError> {
    if let Ok(positive) = i32::try_from(value) {
        return Ok(positive);
    }
    if value < MIN_SIGN_EXTENDED_I32 {
        return Err(DecodeError::invalid());
    }
    i32::try_from(i64::from_ne_bytes(value.to_ne_bytes()))
        .map_err(|_conversion| DecodeError::invalid())
}

fn canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::invalid()),
    }
}

struct Budget {
    options: DecodeOptions,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    reference_bytes: usize,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        let hard_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
            .map_err(|_conversion| DecodeError::invalid())?;
        if options.max_message_bytes > hard_bytes {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: options.max_message_bytes,
                maximum: hard_bytes,
            }));
        }
        if source.len() > options.max_message_bytes {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }));
        }
        if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION,
            }));
        }
        let mut budget = Self {
            options,
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            references: 0,
            reference_bytes: 0,
        };
        budget.observe_depth(1)?;
        Ok(budget)
    }

    fn charge_field(&mut self) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(1);
        if observed > self.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed,
                maximum: self.options.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_reference(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self.references.saturating_add(1);
        if observed > self.options.max_references {
            return Err(DecodeError::limited(DecodeLimit::References {
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

    fn charge_work(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self.work_bytes.saturating_add(amount);
        if observed > self.options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    const fn report(&self) -> DecodeReport {
        DecodeReport {
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            reference_bytes: self.reference_bytes,
        }
    }
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "Focused wire fixtures require exact successful construction and failures."
)]
mod tests {
    use super::*;

    fn push_varint(output: &mut Vec<u8>, mut value: u64) {
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

    fn push_key(output: &mut Vec<u8>, number: u32, wire_type: u8) {
        push_varint(output, (u64::from(number) << 3) | u64::from(wire_type));
    }

    fn push_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
        push_key(output, number, 0);
        push_varint(output, value);
    }

    fn push_fixed64_field(output: &mut Vec<u8>, number: u32, value: u64) {
        push_key(output, number, 1);
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn push_length_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
        push_key(output, number, 2);
        push_varint(output, u64::try_from(payload.len()).unwrap());
        output.extend_from_slice(payload);
    }

    fn rich_reference(identifier: u64) -> Vec<u8> {
        let mut reference = Vec::new();
        push_varint_field(&mut reference, 1, identifier);
        push_varint_field(&mut reference, 2, u64::MAX);
        push_varint_field(&mut reference, 3, 0);
        reference
    }

    fn title_payload(
        visible: Option<bool>,
        height_bits: Option<u64>,
        outlined: Option<bool>,
        paragraph: Option<&[u8]>,
        shape: Option<&[u8]>,
    ) -> Vec<u8> {
        let mut source = Vec::new();
        if let Some(value) = visible {
            push_varint_field(&mut source, TABLE_NAME_ENABLED_FIELD, u64::from(value));
        }
        if let Some(reference) = paragraph {
            push_length_field(&mut source, TABLE_NAME_STYLE_FIELD, reference);
        }
        if let Some(bits) = height_bits {
            push_fixed64_field(&mut source, TABLE_NAME_HEIGHT_FIELD, bits);
        }
        if let Some(reference) = shape {
            push_length_field(&mut source, TABLE_NAME_SHAPE_STYLE_FIELD, reference);
        }
        if let Some(value) = outlined {
            push_varint_field(
                &mut source,
                TABLE_NAME_BORDER_ENABLED_FIELD,
                u64::from(value),
            );
        }
        source
    }

    fn generous(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(source.len().max(1), usize::MAX, usize::MAX, 64, 2)
    }

    fn decode(source: &[u8]) -> Result<TableTitleSettingsSnapshot, DecodeError> {
        decode_table_title_settings(source, generous(source))
    }

    #[test]
    fn full_snapshot_forces_both_private_lazy_views_and_reports_exact_work() {
        let paragraph = rich_reference(41);
        let shape = rich_reference(42);
        let source = title_payload(
            Some(true),
            Some((-0.0f64).to_bits()),
            Some(false),
            Some(&paragraph),
            Some(&shape),
        );
        let (snapshot, report) =
            decode_table_title_settings_with_report(&source, generous(&source)).unwrap();
        assert_eq!(snapshot.table_name_enabled(), Some(true));
        assert_eq!(snapshot.table_name_height_bits(), Some((-0.0f64).to_bits()));
        assert_eq!(snapshot.table_name_border_enabled(), Some(false));
        assert_eq!(
            snapshot
                .table_name_style()
                .map(ReferenceSnapshot::identifier),
            Some(41)
        );
        assert_eq!(
            snapshot
                .table_name_shape_style()
                .map(ReferenceSnapshot::identifier),
            Some(42)
        );
        assert_eq!(
            snapshot
                .table_name_style()
                .and_then(ReferenceSnapshot::deprecated_type),
            Some(-1)
        );
        assert_eq!(report.fields(), 11);
        assert_eq!(report.references(), 2);
        assert_eq!(report.reference_bytes(), paragraph.len() + shape.len());
        assert_eq!(
            report.work_bytes(),
            source.len() * 2 + report.reference_bytes() * 2
        );
        assert_eq!(report.max_depth(), 2);
    }

    #[test]
    fn scalar_presence_and_all_ieee754_bits_are_lossless() {
        let height_bits = [
            0.0f64.to_bits(),
            (-0.0f64).to_bits(),
            f64::NAN.to_bits(),
            f64::INFINITY.to_bits(),
            f64::NEG_INFINITY.to_bits(),
        ];
        for visible in [None, Some(false), Some(true)] {
            for outlined in [None, Some(false), Some(true)] {
                for height in [None].into_iter().chain(height_bits.into_iter().map(Some)) {
                    let source = title_payload(visible, height, outlined, None, None);
                    let snapshot = decode(&source).unwrap();
                    assert_eq!(snapshot.table_name_enabled(), visible);
                    assert_eq!(snapshot.table_name_height_bits(), height);
                    assert_eq!(snapshot.table_name_border_enabled(), outlined);
                    assert_eq!(snapshot.table_name_style(), None);
                    assert_eq!(snapshot.table_name_shape_style(), None);
                }
            }
        }
    }

    #[test]
    fn exact_limits_are_inclusive_and_max_minus_one_is_typed() {
        let paragraph = rich_reference(11);
        let shape = rich_reference(12);
        let source = title_payload(
            Some(true),
            Some(12.5f64.to_bits()),
            Some(true),
            Some(&paragraph),
            Some(&shape),
        );
        let (_, exact) =
            decode_table_title_settings_with_report(&source, generous(&source)).unwrap();
        let exact_options = DecodeOptions::new(
            source.len(),
            exact.fields(),
            exact.work_bytes(),
            exact.max_depth(),
            exact.references(),
        );
        assert!(decode_table_title_settings(&source, exact_options).is_ok());

        let bytes = decode_table_title_settings(
            &source,
            DecodeOptions::new(
                source.len() - 1,
                exact.fields(),
                exact.work_bytes(),
                exact.max_depth(),
                exact.references(),
            ),
        )
        .unwrap_err();
        assert_eq!(
            bytes.resource_limit(),
            Some(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1,
            })
        );

        let fields = decode_table_title_settings(
            &source,
            DecodeOptions::new(
                source.len(),
                exact.fields() - 1,
                exact.work_bytes(),
                exact.max_depth(),
                exact.references(),
            ),
        )
        .unwrap_err();
        assert_eq!(
            fields.resource_limit(),
            Some(DecodeLimit::Fields {
                observed: exact.fields(),
                maximum: exact.fields() - 1,
            })
        );

        let work = decode_table_title_settings(
            &source,
            DecodeOptions::new(
                source.len(),
                exact.fields(),
                exact.work_bytes() - 1,
                exact.max_depth(),
                exact.references(),
            ),
        )
        .unwrap_err();
        assert_eq!(
            work.resource_limit(),
            Some(DecodeLimit::Work {
                observed: exact.work_bytes(),
                maximum: exact.work_bytes() - 1,
            })
        );

        let references = decode_table_title_settings(
            &source,
            DecodeOptions::new(
                source.len(),
                exact.fields(),
                exact.work_bytes(),
                exact.max_depth(),
                exact.references() - 1,
            ),
        )
        .unwrap_err();
        assert_eq!(
            references.resource_limit(),
            Some(DecodeLimit::References {
                observed: exact.references(),
                maximum: exact.references() - 1,
            })
        );

        let nesting = decode_table_title_settings(
            &source,
            DecodeOptions::new(
                source.len(),
                exact.fields(),
                exact.work_bytes(),
                exact.max_depth() - 1,
                exact.references(),
            ),
        )
        .unwrap_err();
        assert_eq!(
            nesting.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: exact.max_depth(),
                maximum: exact.max_depth() - 1,
            })
        );
    }

    #[test]
    fn invalid_configured_nesting_and_buffa_byte_ceiling_are_typed() {
        let zero = decode_table_title_settings(&[], DecodeOptions::new(1, 1, 1, 0, 0)).unwrap_err();
        assert_eq!(
            zero.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 0,
                maximum: MAX_RECURSION,
            })
        );
        let excessive =
            decode_table_title_settings(&[], DecodeOptions::new(1, 1, 1, MAX_RECURSION + 1, 0))
                .unwrap_err();
        assert_eq!(
            excessive.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: MAX_RECURSION + 1,
                maximum: MAX_RECURSION,
            })
        );
        let hard = usize::try_from(buffa::MAX_MESSAGE_BYTES).unwrap();
        let bytes =
            decode_table_title_settings(&[], DecodeOptions::new(hard + 1, 1, 1, 1, 0)).unwrap_err();
        assert_eq!(
            bytes.resource_limit(),
            Some(DecodeLimit::Bytes {
                observed: hard + 1,
                maximum: hard,
            })
        );
    }

    #[test]
    fn selected_fields_reject_duplicates_wrong_wires_and_noncanonical_values() {
        let reference = rich_reference(9);
        for field in [
            TABLE_NAME_ENABLED_FIELD,
            TABLE_NAME_STYLE_FIELD,
            TABLE_NAME_HEIGHT_FIELD,
            TABLE_NAME_SHAPE_STYLE_FIELD,
            TABLE_NAME_BORDER_ENABLED_FIELD,
        ] {
            let mut source = Vec::new();
            match field {
                TABLE_NAME_ENABLED_FIELD | TABLE_NAME_BORDER_ENABLED_FIELD => {
                    push_varint_field(&mut source, field, 0);
                    push_varint_field(&mut source, field, 1);
                },
                TABLE_NAME_STYLE_FIELD | TABLE_NAME_SHAPE_STYLE_FIELD => {
                    push_length_field(&mut source, field, &reference);
                    push_length_field(&mut source, field, &reference);
                },
                TABLE_NAME_HEIGHT_FIELD => {
                    push_fixed64_field(&mut source, field, 0);
                    push_fixed64_field(&mut source, field, 1);
                },
                _ => unreachable!(),
            }
            assert!(decode(&source).is_err(), "duplicate field {field}");

            let mut wrong_wire = Vec::new();
            push_varint_field(&mut wrong_wire, field, 0);
            if matches!(
                field,
                TABLE_NAME_STYLE_FIELD | TABLE_NAME_HEIGHT_FIELD | TABLE_NAME_SHAPE_STYLE_FIELD
            ) {
                assert!(decode(&wrong_wire).is_err(), "wrong wire field {field}");
            }
        }

        for field in [TABLE_NAME_ENABLED_FIELD, TABLE_NAME_BORDER_ENABLED_FIELD] {
            let mut source = Vec::new();
            push_varint_field(&mut source, field, 2);
            assert!(decode(&source).is_err());
        }

        assert!(decode(&[0]).is_err());
        assert!(decode(&[0x80, 0]).is_err());
        assert!(decode(&[0x0e]).is_err());
        let mut truncated_fixed = Vec::new();
        push_key(&mut truncated_fixed, TABLE_NAME_HEIGHT_FIELD, 1);
        truncated_fixed.push(0);
        assert!(decode(&truncated_fixed).is_err());
        let mut overlong_value = Vec::new();
        push_key(&mut overlong_value, TABLE_NAME_ENABLED_FIELD, 0);
        overlong_value.extend_from_slice(&[0x80, 0]);
        assert!(decode(&overlong_value).is_err());
        let mut overlong_length = Vec::new();
        push_key(&mut overlong_length, TABLE_NAME_STYLE_FIELD, 2);
        overlong_length.extend_from_slice(&[0x80, 0]);
        assert!(decode(&overlong_length).is_err());
    }

    #[test]
    fn references_reject_missing_duplicate_zero_and_noncanonical_scalars() {
        let mut invalid_references = Vec::new();
        invalid_references.push(Vec::new());
        let mut zero = Vec::new();
        push_varint_field(&mut zero, 1, 0);
        invalid_references.push(zero);
        let mut duplicate = Vec::new();
        push_varint_field(&mut duplicate, 1, 1);
        push_varint_field(&mut duplicate, 1, 2);
        invalid_references.push(duplicate);
        let mut wrong_wire = Vec::new();
        push_fixed64_field(&mut wrong_wire, 1, 1);
        invalid_references.push(wrong_wire);
        let mut bad_int32 = Vec::new();
        push_varint_field(&mut bad_int32, 1, 1);
        push_varint_field(&mut bad_int32, 2, 0x8000_0000);
        invalid_references.push(bad_int32);
        let mut bad_bool = Vec::new();
        push_varint_field(&mut bad_bool, 1, 1);
        push_varint_field(&mut bad_bool, 3, 2);
        invalid_references.push(bad_bool);
        let mut duplicate_type = Vec::new();
        push_varint_field(&mut duplicate_type, 1, 1);
        push_varint_field(&mut duplicate_type, 2, 1);
        push_varint_field(&mut duplicate_type, 2, 1);
        invalid_references.push(duplicate_type);

        for reference in invalid_references {
            let source = title_payload(None, None, None, Some(&reference), None);
            assert!(decode(&source).is_err());
        }

        let valid_negative = rich_reference(7);
        let source = title_payload(None, None, None, Some(&valid_negative), None);
        assert_eq!(
            decode(&source)
                .unwrap()
                .table_name_style()
                .and_then(ReferenceSnapshot::deprecated_type),
            Some(-1)
        );
    }

    #[test]
    fn canonical_unknown_fields_and_matched_groups_are_accepted() {
        let mut source = Vec::new();
        push_varint_field(&mut source, 100, 7);
        push_fixed64_field(&mut source, 101, u64::MAX);
        push_length_field(&mut source, 102, &[0xff, 0x00]);
        push_key(&mut source, 103, 5);
        source.extend_from_slice(&[1, 2, 3, 4]);
        push_key(&mut source, 104, 3);
        push_varint_field(&mut source, 1, 9);
        push_key(&mut source, 105, 3);
        push_varint_field(&mut source, 2, 10);
        push_key(&mut source, 105, 4);
        push_key(&mut source, 104, 4);
        push_varint_field(&mut source, TABLE_NAME_ENABLED_FIELD, 1);
        let (snapshot, report) =
            decode_table_title_settings_with_report(&source, generous(&source)).unwrap();
        assert_eq!(snapshot.table_name_enabled(), Some(true));
        assert_eq!(report.fields(), 11);
        assert_eq!(report.max_depth(), 3);
        assert_eq!(report.references(), 0);
        assert_eq!(report.work_bytes(), source.len() * 2);
    }

    #[test]
    fn malformed_groups_fail_closed() {
        let mut unterminated = Vec::new();
        push_key(&mut unterminated, 100, 3);
        push_varint_field(&mut unterminated, 1, 1);
        assert!(decode(&unterminated).is_err());

        let mut mismatched = Vec::new();
        push_key(&mut mismatched, 100, 3);
        push_key(&mut mismatched, 101, 4);
        assert!(decode(&mismatched).is_err());

        let mut unexpected = Vec::new();
        push_key(&mut unexpected, 100, 4);
        assert!(decode(&unexpected).is_err());

        let mut nested = Vec::new();
        push_key(&mut nested, 100, 3);
        push_key(&mut nested, 101, 3);
        push_key(&mut nested, 101, 4);
        push_key(&mut nested, 100, 4);
        let error = decode_table_title_settings(
            &nested,
            DecodeOptions::new(nested.len(), 4, nested.len() * 2, 2, 0),
        )
        .unwrap_err();
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 3,
                maximum: 2,
            })
        );
    }

    #[test]
    fn wide_unknown_field_routing_scales_linearly_and_limits_before_crosscheck() {
        fn wide(fields: usize) -> Vec<u8> {
            let mut source = Vec::new();
            for index in 0..fields {
                push_varint_field(&mut source, 100, u64::try_from(index).unwrap());
            }
            source
        }

        const SMALL: usize = 4_096;
        const LARGE: usize = 8_192;
        let small_source = wide(SMALL);
        let large_source = wide(LARGE);
        let (_, small) = decode_table_title_settings_with_report(
            &small_source,
            DecodeOptions::new(small_source.len(), SMALL, usize::MAX, 1, 0),
        )
        .unwrap();
        let (_, large) = decode_table_title_settings_with_report(
            &large_source,
            DecodeOptions::new(large_source.len(), LARGE, usize::MAX, 1, 0),
        )
        .unwrap();
        assert_eq!(small.fields(), SMALL);
        assert_eq!(large.fields(), LARGE);
        assert_eq!(small.work_bytes(), small_source.len() * 2);
        assert_eq!(large.work_bytes(), large_source.len() * 2);
        assert!(large_source.len() * 10 <= small_source.len() * 23);
        assert!(large.fields() * 10 <= small.fields() * 23);
        assert!(large.work_bytes() * 10 <= small.work_bytes() * 23);

        let field_error = decode_table_title_settings(
            &large_source,
            DecodeOptions::new(large_source.len(), LARGE - 1, large.work_bytes(), 1, 0),
        )
        .unwrap_err();
        assert_eq!(
            field_error.resource_limit(),
            Some(DecodeLimit::Fields {
                observed: LARGE,
                maximum: LARGE - 1,
            })
        );
        let work_error = decode_table_title_settings(
            &large_source,
            DecodeOptions::new(large_source.len(), LARGE, large.work_bytes() - 1, 1, 0),
        )
        .unwrap_err();
        assert_eq!(
            work_error.resource_limit(),
            Some(DecodeLimit::Work {
                observed: large.work_bytes(),
                maximum: large.work_bytes() - 1,
            })
        );
    }
}
