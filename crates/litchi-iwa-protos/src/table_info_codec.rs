//! Strict private Buffa projection for `TST.TableInfoArchive.table_model`.
//!
//! The strict raw-wire pass canonicalizes every visited field framing before
//! Buffa observes the source. It selects only the required model reference and
//! its non-zero identifier; `super` and all unselected table metadata remain
//! opaque caller-owned bytes and are never materialized or re-encoded.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict preflight intentionally precedes the low-level wire reader it consumes."
)]

use std::{fmt, num::NonZeroU64};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_table_info_generated::LitchiIwaProjection as projection;

const TABLE_SUPER_FIELD: u32 = 1;
const TABLE_MODEL_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const MAX_RECURSION_LIMIT: u32 = 64;

/// Explicit finite resource policy for one `TableInfo` model-reference decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build a finite bytes/fields/work/nesting policy.
    ///
    /// Work accounts for both the strict scan and Buffa's deferred model
    /// reference access. The selected nested reference is charged separately
    /// because it is scanned and forced after the outer message.
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

    fn descend(self) -> Result<Self, DecodeError> {
        let recursion_limit = self
            .recursion_limit
            .checked_sub(1)
            .ok_or_else(DecodeError::recursion_limit)?;
        Ok(Self {
            recursion_limit,
            ..self
        })
    }
}

/// Typed non-zero reference to a native `TST.TableModelArchive` object.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableModelReference {
    identifier: NonZeroU64,
}

impl TableModelReference {
    /// Native model object identifier, proven non-zero by strict preflight.
    #[must_use]
    pub const fn identifier(self) -> NonZeroU64 {
        self.identifier
    }
}

/// Failure from `TableInfo` strict preflight or its Buffa cross-check.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    ZeroIdentifier(&'static str),
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
    Projection,
}

impl DecodeError {
    fn recursion_limit() -> Self {
        buffa::DecodeError::RecursionLimitExceeded.into()
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

    const fn zero_identifier(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::ZeroIdentifier(field),
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

    /// Required schema field absent from the source, when applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Singular schema field repeated in the source, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
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
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Reference field carrying a forbidden zero identifier, when applicable.
    #[must_use]
    pub const fn zero_identifier_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::ZeroIdentifier(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Observed and configured field counts for a field-limit failure.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::FieldLimit { observed, maximum } => Some((observed, maximum)),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Observed and configured work bytes for a work-limit failure.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::WorkLimit { observed, maximum } => Some((observed, maximum)),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::ZeroIdentifier(field) => write!(formatter, "{field} is zero"),
            DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Numbers TableInfo projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Numbers TableInfo projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Projection => formatter.write_str(
                "Numbers TableInfo strict preflight disagrees with the Buffa projection",
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

/// Decode the required `TableInfo` table-model reference from preflighted bytes.
///
/// Every root field is strictly scanned for canonical protobuf framing. The
/// required opaque `super` envelope is checked only for unique
/// length-delimited framing; field 2's deferred reference is then forced once
/// after strict preflight. Both `super` and all unselected metadata remain
/// opaque in the caller-owned source representation.
pub fn decode_table_model_reference(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableModelReference, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_table_info(source, options, &mut budget)?;

    let view: projection::TableInfoArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    if !view.has_table_model() {
        return Err(DecodeError::missing_required(
            "TST.TableInfoArchive.table_model",
        ));
    }
    let model = view
        .table_model
        .get()?
        .ok_or_else(|| DecodeError::missing_required("TST.TableInfoArchive.table_model"))?;
    if !model.has_identifier() {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    let projected = TableModelReference {
        identifier: NonZeroU64::new(model.identifier)
            .ok_or_else(|| DecodeError::zero_identifier("TSP.Reference.identifier"))?,
    };
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let max_buffa_message_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
    if options.max_message_bytes > max_buffa_message_bytes
        || source.len() > options.max_message_bytes
    {
        return Err(buffa::DecodeError::MessageTooLarge.into());
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::recursion_limit());
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

fn preflight_table_info(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<TableModelReference, DecodeError> {
    budget.charge_message(source.len())?;
    let nested_options = options.descend()?;
    let mut model = None;
    let mut saw_super = false;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            TABLE_SUPER_FIELD => {
                if saw_super {
                    return Err(DecodeError::duplicate_singular(
                        "TST.TableInfoArchive.super",
                    ));
                }
                saw_super = true;
                let _opaque_super = field.length_delimited()?;
            },
            TABLE_MODEL_FIELD => {
                if model.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TST.TableInfoArchive.table_model",
                    ));
                }
                model = Some(preflight_reference(
                    field.length_delimited()?,
                    nested_options,
                    budget,
                )?);
            },
            _ => {},
        }
    }
    if !saw_super {
        return Err(DecodeError::missing_required("TST.TableInfoArchive.super"));
    }
    model.ok_or_else(|| DecodeError::missing_required("TST.TableInfoArchive.table_model"))
}

fn preflight_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<TableModelReference, DecodeError> {
    budget.charge_message(source.len())?;
    let mut identifier = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        if field.number != REFERENCE_IDENTIFIER_FIELD {
            continue;
        }
        if identifier.is_some() {
            return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
        }
        identifier = Some(
            NonZeroU64::new(field.varint()?)
                .ok_or_else(|| DecodeError::zero_identifier("TSP.Reference.identifier"))?,
        );
    }
    Ok(TableModelReference {
        identifier: identifier
            .ok_or_else(|| DecodeError::missing_required("TSP.Reference.identifier"))?,
    })
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

#[derive(Clone, Copy, Debug)]
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
            let _bytes = take_exact(source, 8)?;
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
            let child_limit = recursion_limit
                .checked_sub(1)
                .ok_or_else(DecodeError::recursion_limit)?;
            skip_strict_group(source, field_number, child_limit, budget)?;
            StrictValue::Group
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            let _bytes = take_exact(source, 4)?;
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
    reason = "Focused negative tests use explicit panic messages and reuse local error roles."
)]
mod tests {
    use std::num::NonZeroU64;

    use prost::Message as _;

    use super::{DecodeOptions, TableModelReference, decode_table_model_reference};
    use crate::{tsd, tsp, tst};

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            source.len().max(1),
            source.len().saturating_mul(4).max(1),
            2,
        )
    }

    fn decode(source: &[u8]) -> Result<TableModelReference, super::DecodeError> {
        decode_table_model_reference(source, options(source))
    }

    fn table_model_field(model: &[u8]) -> Vec<u8> {
        let mut output = vec![0x12, u8::try_from(model.len()).expect("small test payload")];
        output.extend_from_slice(model);
        output
    }

    fn table_info(model: &[u8]) -> Vec<u8> {
        [vec![0x0a, 0x00], table_model_field(model)].concat()
    }

    #[test]
    fn canonical_prost_table_info_matches_the_strict_projection()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = tst::TableInfoArchive {
            super_: tsd::DrawableArchive::default(),
            table_model: tsp::Reference {
                identifier: 42,
                deprecated_type: Some(-7),
                deprecated_is_external: Some(false),
            },
            ..tst::TableInfoArchive::default()
        }
        .encode_to_vec();

        assert_eq!(
            decode(&source)?.identifier(),
            NonZeroU64::new(42).expect("non-zero test identifier")
        );
        Ok(())
    }

    #[test]
    fn opaque_super_and_unselected_metadata_are_not_decoded()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = [
            0x0a, 0x01, 0xff, // opaque DrawableArchive super
            0x12, 0x02, 0x08, 0x2a, // selected table-model reference
            0x1a, 0x01, 0xff, // opaque editing_state metadata
        ];
        assert_eq!(
            decode(&source)?.identifier(),
            NonZeroU64::new(42).expect("non-zero test identifier")
        );
        Ok(())
    }

    #[test]
    fn required_and_unique_table_model_identifier_are_enforced() {
        let error = decode(&[]).expect_err("missing opaque super envelope");
        assert_eq!(
            error.missing_required_field(),
            Some("TST.TableInfoArchive.super")
        );

        let error = decode(&[0x0a, 0x00]).expect_err("missing model reference");
        assert_eq!(
            error.missing_required_field(),
            Some("TST.TableInfoArchive.table_model")
        );

        let error = decode(&table_info(&[])).expect_err("missing nested identifier");
        assert_eq!(
            error.missing_required_field(),
            Some("TSP.Reference.identifier")
        );

        let duplicate_model = [
            vec![0x0a, 0x00],
            table_model_field(&[0x08, 0x01]),
            table_model_field(&[0x08, 0x02]),
        ]
        .concat();
        let error = decode(&duplicate_model).expect_err("duplicate model reference");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TST.TableInfoArchive.table_model")
        );

        let error = decode(&table_info(&[0x08, 0x01, 0x08, 0x02]))
            .expect_err("duplicate nested identifier");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSP.Reference.identifier")
        );

        let duplicate_super = [
            vec![0x0a, 0x00],
            vec![0x0a, 0x00],
            table_model_field(&[0x08, 0x01]),
        ]
        .concat();
        let error = decode(&duplicate_super).expect_err("duplicate opaque super envelope");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TST.TableInfoArchive.super")
        );
    }

    #[test]
    fn wrong_wire_zero_and_malformed_selected_fields_are_rejected() {
        let root_wrong_wire = [0x0a, 0x00, 0x10, 0x2a];
        assert!(decode(&root_wrong_wire).is_err());

        let super_wrong_wire = [0x08, 0x00, 0x12, 0x02, 0x08, 0x2a];
        assert!(decode(&super_wrong_wire).is_err());

        let nested_wrong_wire = table_info(&[0x0a, 0x00]);
        assert!(decode(&nested_wrong_wire).is_err());

        let error = decode(&table_info(&[0x08, 0x00])).expect_err("zero identifier");
        assert_eq!(
            error.zero_identifier_field(),
            Some("TSP.Reference.identifier")
        );

        let malformed = [0x0a, 0x00, 0x12, 0x01, 0x08];
        assert!(decode(&malformed).is_err());
    }

    #[test]
    fn noncanonical_selected_and_opaque_framing_is_rejected() {
        let overlong_super_length = [0x0a, 0x80, 0x00, 0x12, 0x02, 0x08, 0x2a];
        assert_eq!(
            decode(&overlong_super_length)
                .expect_err("overlong opaque super length")
                .noncanonical_reason(),
            Some("length-delimited size")
        );

        let overlong_key = [0x0a, 0x00, 0x92, 0x00, 0x02, 0x08, 0x2a];
        assert_eq!(
            decode(&overlong_key)
                .expect_err("overlong root key")
                .noncanonical_reason(),
            Some("protobuf field key")
        );

        let overlong_length = [0x0a, 0x00, 0x12, 0x82, 0x00, 0x08, 0x2a];
        assert_eq!(
            decode(&overlong_length)
                .expect_err("overlong nested length")
                .noncanonical_reason(),
            Some("length-delimited size")
        );

        let overlong_identifier = table_info(&[0x08, 0xaa, 0x00]);
        assert_eq!(
            decode(&overlong_identifier)
                .expect_err("overlong identifier")
                .noncanonical_reason(),
            Some("protobuf varint value")
        );

        let noncanonical_opaque = [0x0a, 0x00, 0x18, 0x81, 0x00, 0x12, 0x02, 0x08, 0x2a];
        assert_eq!(
            decode(&noncanonical_opaque)
                .expect_err("noncanonical opaque metadata")
                .noncanonical_reason(),
            Some("protobuf varint value")
        );
    }

    #[test]
    fn exact_boundary_limits_are_accepted_and_one_less_is_rejected()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = [0x0a, 0x00, 0x12, 0x02, 0x08, 0x01];
        let exact = DecodeOptions::new(source.len(), 3, 16, 2);
        assert_eq!(
            decode_table_model_reference(&source, exact)?.identifier(),
            NonZeroU64::new(1).expect("non-zero test identifier")
        );
        assert!(
            decode_table_model_reference(&source, DecodeOptions::new(source.len() - 1, 3, 16, 2))
                .is_err()
        );
        assert_eq!(
            decode_table_model_reference(&source, DecodeOptions::new(source.len(), 2, 16, 2))
                .expect_err("field cap")
                .field_limit_values(),
            Some((3, 2))
        );
        assert_eq!(
            decode_table_model_reference(&source, DecodeOptions::new(source.len(), 3, 15, 2))
                .expect_err("work cap")
                .work_limit_values(),
            Some((16, 15))
        );
        assert!(
            decode_table_model_reference(&source, DecodeOptions::new(source.len(), 3, 16, 0))
                .is_err()
        );
        Ok(())
    }

    #[test]
    fn structural_wire_failures_never_panic() {
        let malformed: [&[u8]; 6] = [
            &[0x80],
            &[0x00],
            &[0x0f],
            &[0x12, 0x02, 0x08],
            &[0x0b],
            &[
                0x08, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x02,
            ],
        ];
        for source in malformed {
            assert!(decode(source).is_err());
        }
    }
}
