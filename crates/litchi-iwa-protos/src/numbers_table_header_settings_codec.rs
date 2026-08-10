//! Strict scalar projection for Numbers table header/footer settings.
use crate::buffa_numbers_table_header_settings_generated::LitchiIwaProjection as projection;
use buffa::DecodeOptions as BuffaDecodeOptions;
use std::fmt;
const TABLE_ROWS_FIELD: u32 = 6;
const TABLE_COLUMNS_FIELD: u32 = 7;
const HEADER_ROWS_FIELD: u32 = 9;
const HEADER_COLUMNS_FIELD: u32 = 10;
const FOOTER_ROWS_FIELD: u32 = 11;
const HEADER_ROWS_FROZEN_FIELD: u32 = 12;
const HEADER_COLUMNS_FROZEN_FIELD: u32 = 13;
const REPEATING_HEADER_ROWS_FIELD: u32 = 29;
const REPEATING_HEADER_COLUMNS_FIELD: u32 = 32;
const MAX_RECURSION: u32 = 64;
/// Finite aggregate limits for one table-model payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    bytes: usize,
    fields: usize,
    work: usize,
    recursion: u32,
}
impl DecodeOptions {
    #[must_use]
    pub const fn new(bytes: usize, fields: usize, work: usize, recursion: u32) -> Self {
        Self {
            bytes,
            fields,
            work,
            recursion,
        }
    }
    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion)
    }
}
/// Presence-preserving native header/footer settings; booleans are raw source facts.
/// Content-free byte or nesting limit observation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableHeaderSettingsSnapshot {
    rows: u32,
    columns: u32,
    header_rows: Option<u32>,
    header_columns: Option<u32>,
    footer_rows: Option<u32>,
    header_rows_frozen: Option<bool>,
    header_columns_frozen: Option<bool>,
    repeating_header_rows_enabled: Option<bool>,
    repeating_header_columns_enabled: Option<bool>,
}
impl TableHeaderSettingsSnapshot {
    #[must_use]
    pub const fn rows(self) -> u32 {
        self.rows
    }
    #[must_use]
    pub const fn columns(self) -> u32 {
        self.columns
    }
    #[must_use]
    pub const fn header_rows(self) -> Option<u32> {
        self.header_rows
    }
    #[must_use]
    pub const fn header_columns(self) -> Option<u32> {
        self.header_columns
    }
    #[must_use]
    pub const fn footer_rows(self) -> Option<u32> {
        self.footer_rows
    }
    #[must_use]
    pub const fn header_rows_frozen(self) -> Option<bool> {
        self.header_rows_frozen
    }
    #[must_use]
    pub const fn header_columns_frozen(self) -> Option<bool> {
        self.header_columns_frozen
    }
    #[must_use]
    pub const fn repeating_header_rows_enabled(self) -> Option<bool> {
        self.repeating_header_rows_enabled
    }
    #[must_use]
    pub const fn repeating_header_columns_enabled(self) -> Option<bool> {
        self.repeating_header_columns_enabled
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    Bytes { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
}
/// Strict wire-preflight or private projection failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError(Kind);
#[derive(Debug, Clone, PartialEq, Eq)]
enum Kind {
    Wire(buffa::DecodeError),
    Resource(WireResourceLimit),
    Missing(&'static str),
    Duplicate(&'static str),
    NonCanonical(&'static str),
    Field { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Projection,
}
impl DecodeError {
    const fn resource(x: WireResourceLimit) -> Self {
        Self(Kind::Resource(x))
    }
    const fn missing(x: &'static str) -> Self {
        Self(Kind::Missing(x))
    }
    const fn duplicate(x: &'static str) -> Self {
        Self(Kind::Duplicate(x))
    }
    const fn noncanonical(x: &'static str) -> Self {
        Self(Kind::NonCanonical(x))
    }
    #[must_use]
    pub const fn wire_resource_limit(&self) -> Option<WireResourceLimit> {
        if let Kind::Resource(x) = self.0 {
            Some(x)
        } else {
            None
        }
    }
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        if let Kind::Missing(x) = self.0 {
            Some(x)
        } else {
            None
        }
    }
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        if let Kind::Duplicate(x) = self.0 {
            Some(x)
        } else {
            None
        }
    }
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        if let Kind::NonCanonical(x) = self.0 {
            Some(x)
        } else {
            None
        }
    }
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        if let Kind::Field { observed, maximum } = self.0 {
            Some((observed, maximum))
        } else {
            None
        }
    }
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        if let Kind::Work { observed, maximum } = self.0 {
            Some((observed, maximum))
        } else {
            None
        }
    }
}
impl fmt::Display for DecodeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.0 {
            Kind::Wire(x) => x.fmt(f),
            Kind::Resource(_) => f.write_str("Numbers table-header resource limit exceeded"),
            Kind::Missing(x) => write!(f, "missing required field {x}"),
            Kind::Duplicate(x) => write!(f, "duplicate singular field {x}"),
            Kind::NonCanonical(x) => write!(f, "non-canonical protobuf representation: {x}"),
            Kind::Field { observed, maximum } => {
                write!(f, "visited {observed} fields; maximum is {maximum}")
            },
            Kind::Work { observed, maximum } => {
                write!(f, "requires {observed} work bytes; maximum is {maximum}")
            },
            Kind::Projection => f.write_str("strict preflight disagrees with Buffa projection"),
        }
    }
}
impl std::error::Error for DecodeError {}
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "Buffa errors are non-exhaustive and unrecognized failures remain opaque wire errors."
)]
impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        match error {
            buffa::DecodeError::MessageTooLarge | buffa::DecodeError::RecursionLimitExceeded => {
                Self(Kind::Projection)
            },
            other => Self(Kind::Wire(other)),
        }
    }
}

struct Budget {
    fields: usize,
    work: usize,
    max_fields: usize,
    max_work: usize,
    max_recursion: u32,
}
impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work: 0,
            max_fields: options.fields,
            max_work: options.work,
            max_recursion: options.recursion,
        }
    }
    fn charge(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = bytes
            .checked_mul(2)
            .and_then(|cost| self.work.checked_add(cost))
            .ok_or(DecodeError(Kind::Projection))?;
        if observed > self.max_work {
            return Err(DecodeError(Kind::Work {
                observed,
                maximum: self.max_work,
            }));
        }
        self.work = observed;
        Ok(())
    }
    fn field(&mut self) -> Result<(), DecodeError> {
        let observed = self
            .fields
            .checked_add(1)
            .ok_or(DecodeError(Kind::Projection))?;
        if observed > self.max_fields {
            return Err(DecodeError(Kind::Field {
                observed,
                maximum: self.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }
    const fn nesting(&self) -> DecodeError {
        DecodeError::resource(WireResourceLimit::Nesting {
            observed: self.max_recursion.saturating_add(1),
            maximum: self.max_recursion,
        })
    }
}
/// Decode one `TST.TableModelArchive` header/footer scalar envelope without retaining raw IDs.
pub fn decode_table_header_settings(
    source: &[u8],
    o: DecodeOptions,
) -> Result<TableHeaderSettingsSnapshot, DecodeError> {
    validate(source, o)?;
    let mut b = Budget::new(o);
    let strict = preflight(source, o, &mut b)?;
    let v: projection::NumbersTableHeaderSettingsArchiveLazyView<'_> =
        o.buffa().decode_lazy_view(source)?;
    let projected = TableHeaderSettingsSnapshot {
        rows: v.number_of_rows,
        columns: v.number_of_columns,
        header_rows: v.number_of_header_rows,
        header_columns: v.number_of_header_columns,
        footer_rows: v.number_of_footer_rows,
        header_rows_frozen: v.header_rows_frozen,
        header_columns_frozen: v.header_columns_frozen,
        repeating_header_rows_enabled: v.repeating_header_rows_enabled,
        repeating_header_columns_enabled: v.repeating_header_columns_enabled,
    };
    if projected != strict {
        return Err(DecodeError(Kind::Projection));
    }
    Ok(strict)
}
fn validate(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| DecodeError(Kind::Projection))?;
    if options.bytes > hard {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: options.bytes,
            maximum: hard,
        }));
    }
    if source.len() > options.bytes {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: source.len(),
            maximum: options.bytes,
        }));
    }
    if options.recursion == 0 || options.recursion > MAX_RECURSION {
        return Err(DecodeError::resource(WireResourceLimit::Nesting {
            observed: options.recursion,
            maximum: MAX_RECURSION,
        }));
    }
    Ok(())
}
fn preflight(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<TableHeaderSettingsSnapshot, DecodeError> {
    budget.charge(source.len())?;
    let mut snapshot = TableHeaderSettingsSnapshot {
        rows: 0,
        columns: 0,
        header_rows: None,
        header_columns: None,
        footer_rows: None,
        header_rows_frozen: None,
        header_columns_frozen: None,
        repeating_header_rows_enabled: None,
        repeating_header_columns_enabled: None,
    };
    let mut seen = 0u64;
    let mut remaining = source;
    while !remaining.is_empty() {
        let (tag, key_canonical) = varint(&mut remaining)?;
        if !key_canonical {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
        budget.field()?;
        let raw =
            u32::try_from(tag).map_err(|_conversion| buffa::DecodeError::InvalidFieldNumber)?;
        let field_number = raw >> 3;
        if field_number == 0 {
            return Err(buffa::DecodeError::InvalidFieldNumber.into());
        }
        if raw & 7 != 0 {
            skip(
                &mut remaining,
                field_number,
                raw & 7,
                options.recursion,
                budget,
            )?;
            if matches!(
                field_number,
                TABLE_ROWS_FIELD
                    | TABLE_COLUMNS_FIELD
                    | HEADER_ROWS_FIELD
                    | HEADER_COLUMNS_FIELD
                    | FOOTER_ROWS_FIELD
                    | HEADER_ROWS_FROZEN_FIELD
                    | HEADER_COLUMNS_FROZEN_FIELD
                    | REPEATING_HEADER_ROWS_FIELD
                    | REPEATING_HEADER_COLUMNS_FIELD
            ) {
                return Err(buffa::DecodeError::WireTypeMismatch {
                    field_number,
                    expected: 0,
                    actual: (raw & 7) as u8,
                }
                .into());
            }
            continue;
        }
        let (value, value_canonical) = varint(&mut remaining)?;
        if !value_canonical {
            return Err(DecodeError::noncanonical("protobuf varint value"));
        }
        let bit = 1u64.checked_shl(field_number).unwrap_or(0);
        let name = match field_number {
            TABLE_ROWS_FIELD => "TST.TableModelArchive.number_of_rows",
            TABLE_COLUMNS_FIELD => "TST.TableModelArchive.number_of_columns",
            HEADER_ROWS_FIELD => "TST.TableModelArchive.number_of_header_rows",
            HEADER_COLUMNS_FIELD => "TST.TableModelArchive.number_of_header_columns",
            FOOTER_ROWS_FIELD => "TST.TableModelArchive.number_of_footer_rows",
            HEADER_ROWS_FROZEN_FIELD => "TST.TableModelArchive.header_rows_frozen",
            HEADER_COLUMNS_FROZEN_FIELD => "TST.TableModelArchive.header_columns_frozen",
            REPEATING_HEADER_ROWS_FIELD => "TST.TableModelArchive.repeating_header_rows_enabled",
            REPEATING_HEADER_COLUMNS_FIELD => {
                "TST.TableModelArchive.repeating_header_columns_enabled"
            },
            _ => continue,
        };
        if seen & bit != 0 {
            return Err(DecodeError::duplicate(name));
        }
        seen |= bit;
        let parsed_u32 = u32::try_from(value)
            .map_err(|_conversion| DecodeError::noncanonical("uint32 scalar exceeds u32"))?;
        match field_number {
            TABLE_ROWS_FIELD => snapshot.rows = parsed_u32,
            TABLE_COLUMNS_FIELD => snapshot.columns = parsed_u32,
            HEADER_ROWS_FIELD => snapshot.header_rows = Some(parsed_u32),
            HEADER_COLUMNS_FIELD => snapshot.header_columns = Some(parsed_u32),
            FOOTER_ROWS_FIELD => snapshot.footer_rows = Some(parsed_u32),
            HEADER_ROWS_FROZEN_FIELD => snapshot.header_rows_frozen = Some(boolean(value)?),
            HEADER_COLUMNS_FROZEN_FIELD => snapshot.header_columns_frozen = Some(boolean(value)?),
            REPEATING_HEADER_ROWS_FIELD => {
                snapshot.repeating_header_rows_enabled = Some(boolean(value)?);
            },
            REPEATING_HEADER_COLUMNS_FIELD => {
                snapshot.repeating_header_columns_enabled = Some(boolean(value)?);
            },
            _ => {},
        }
    }
    if seen & (1 << TABLE_ROWS_FIELD) == 0 {
        return Err(DecodeError::missing("TST.TableModelArchive.number_of_rows"));
    }
    if seen & (1 << TABLE_COLUMNS_FIELD) == 0 {
        return Err(DecodeError::missing(
            "TST.TableModelArchive.number_of_columns",
        ));
    }
    Ok(snapshot)
}
fn boolean(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}
fn varint(source: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *source;
    let mut value = 0;
    for index in 0..10 {
        let byte = *original
            .get(index)
            .ok_or(buffa::DecodeError::UnexpectedEof)?;
        if index == 9 && byte > 1 {
            return Err(buffa::DecodeError::VarintTooLong.into());
        }
        value |= u64::from(byte & 127) << (index * 7);
        if byte & 128 == 0 {
            *source = &original[index + 1..];
            let mut remaining = value;
            let mut canonical_length = 1;
            while remaining >= 128 {
                remaining >>= 7;
                canonical_length += 1;
            }
            return Ok((value, canonical_length == index + 1));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}
fn skip(
    s: &mut &[u8],
    number: u32,
    wire: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    match wire {
        0 => {
            let (_, canonical) = varint(s)?;
            if !canonical {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            }
            Ok(())
        },
        1 => take(s, 8),
        2 => {
            let (length, canonical) = varint(s)?;
            if !canonical {
                return Err(DecodeError::noncanonical("length-delimited size"));
            }
            take(
                s,
                usize::try_from(length)
                    .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?,
            )
        },
        3 => skip_group(
            s,
            number,
            depth.checked_sub(1).ok_or_else(|| budget.nesting())?,
            budget,
        ),
        4 => Err(buffa::DecodeError::InvalidEndGroup(number).into()),
        5 => take(s, 4),
        _ => Err(buffa::DecodeError::InvalidWireType(wire).into()),
    }
}

fn skip_group(
    s: &mut &[u8],
    expected: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    loop {
        if s.is_empty() {
            return Err(buffa::DecodeError::UnexpectedEof.into());
        }
        let (tag, canonical) = varint(s)?;
        if !canonical {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
        budget.field()?;
        let raw =
            u32::try_from(tag).map_err(|_conversion| buffa::DecodeError::InvalidFieldNumber)?;
        let number = raw >> 3;
        if number == 0 {
            return Err(buffa::DecodeError::InvalidFieldNumber.into());
        }
        let wire = raw & 7;
        if wire == 4 {
            if number == expected {
                return Ok(());
            }
            return Err(buffa::DecodeError::InvalidEndGroup(number).into());
        }
        skip(s, number, wire, depth, budget)?;
    }
}
fn take(s: &mut &[u8], n: usize) -> Result<(), DecodeError> {
    if s.len() < n {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    *s = &s[n..];
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn push_varint(mut value: u64, output: &mut Vec<u8>) {
        while value >= 128 {
            output.push((value as u8 & 127) | 128);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(u64::from(number) << 3, &mut output);
        push_varint(value, &mut output);
        output
    }

    fn dimensions() -> Vec<u8> {
        let mut output = varint_field(TABLE_ROWS_FIELD, 1);
        output.extend(varint_field(TABLE_COLUMNS_FIELD, 1));
        output
    }

    #[test]
    fn preserves_header_presence_and_rejects_strict_failures() -> Result<(), DecodeError> {
        let source = [
            0x30, 10, 0x38, 5, 0x48, 2, 0x50, 3, 0x58, 1, 0x60, 1, 0x68, 0, 0xe8, 1, 1, 0x80, 2, 0,
        ];
        let snapshot = decode_table_header_settings(
            &source,
            DecodeOptions::new(source.len(), 9, source.len() * 2, 2),
        )?;
        assert_eq!(snapshot.rows(), 10);
        assert_eq!(snapshot.columns(), 5);
        assert_eq!(snapshot.header_rows(), Some(2));
        assert_eq!(snapshot.header_columns(), Some(3));
        assert_eq!(snapshot.footer_rows(), Some(1));
        assert_eq!(snapshot.header_rows_frozen(), Some(true));
        assert_eq!(snapshot.header_columns_frozen(), Some(false));
        assert_eq!(snapshot.repeating_header_rows_enabled(), Some(true));
        assert_eq!(snapshot.repeating_header_columns_enabled(), Some(false));
        let duplicate = [0x30, 1, 0x30, 2, 0x38, 1];
        assert_eq!(
            decode_table_header_settings(
                &duplicate,
                DecodeOptions::new(duplicate.len(), 3, duplicate.len() * 2, 2)
            )
            .expect_err("duplicate")
            .duplicate_singular_field(),
            Some("TST.TableModelArchive.number_of_rows")
        );
        let bad_bool = [0x30, 1, 0x38, 1, 0x60, 2];
        assert_eq!(
            decode_table_header_settings(
                &bad_bool,
                DecodeOptions::new(bad_bool.len(), 3, bad_bool.len() * 2, 2)
            )
            .expect_err("bool")
            .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
        Ok(())
    }

    #[test]
    fn unknown_groups_are_bounded_and_canonical() -> Result<(), DecodeError> {
        let one_group = [0x30, 1, 0x38, 1, 0xa3, 0x06, 0x08, 1, 0xa4, 0x06];
        assert!(
            decode_table_header_settings(
                &one_group,
                DecodeOptions::new(one_group.len(), 5, one_group.len() * 2, 1)
            )
            .is_ok()
        );
        assert_eq!(
            decode_table_header_settings(
                &one_group,
                DecodeOptions::new(one_group.len(), 4, one_group.len() * 2, 1)
            )
            .expect_err("fields")
            .field_limit_values(),
            Some((5, 4))
        );
        let dimensions = [0x30, 1, 0x38, 1];
        assert_eq!(
            decode_table_header_settings(&dimensions, DecodeOptions::new(4, 2, 7, 1))
                .expect_err("work")
                .work_limit_values(),
            Some((8, 7))
        );
        let nested_groups = [
            0x30, 1, 0x38, 1, 0xa3, 0x06, 0xab, 0x06, 0xac, 0x06, 0xa4, 0x06,
        ];
        assert_eq!(
            decode_table_header_settings(
                &nested_groups,
                DecodeOptions::new(nested_groups.len(), 5, nested_groups.len() * 2, 1)
            )
            .expect_err("depth")
            .wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 2,
                maximum: 1
            })
        );
        let mismatched = [0x30, 1, 0x38, 1, 0xa3, 0x06, 0xac, 0x06];
        assert!(
            decode_table_header_settings(
                &mismatched,
                DecodeOptions::new(mismatched.len(), 4, mismatched.len() * 2, 2)
            )
            .is_err()
        );
        let noncanonical_length = [0x30, 1, 0x38, 1, 0x82, 0x06, 0x80, 0, 0];
        assert_eq!(
            decode_table_header_settings(
                &noncanonical_length,
                DecodeOptions::new(
                    noncanonical_length.len(),
                    3,
                    noncanonical_length.len() * 2,
                    2
                )
            )
            .expect_err("length")
            .noncanonical_reason(),
            Some("length-delimited size")
        );
        Ok(())
    }

    #[test]
    fn table_driven_strict_wire_regressions() {
        let dims = dimensions();
        for (source, required) in [
            (
                varint_field(TABLE_ROWS_FIELD, 1),
                "TST.TableModelArchive.number_of_columns",
            ),
            (
                varint_field(TABLE_COLUMNS_FIELD, 1),
                "TST.TableModelArchive.number_of_rows",
            ),
        ] {
            assert_eq!(
                decode_table_header_settings(
                    &source,
                    DecodeOptions::new(source.len(), 9, source.len() * 2, 2)
                )
                .expect_err("missing dimension")
                .missing_required_field(),
                Some(required)
            );
        }
        for (field, name) in [
            (TABLE_ROWS_FIELD, "TST.TableModelArchive.number_of_rows"),
            (
                TABLE_COLUMNS_FIELD,
                "TST.TableModelArchive.number_of_columns",
            ),
            (
                HEADER_ROWS_FIELD,
                "TST.TableModelArchive.number_of_header_rows",
            ),
            (
                HEADER_COLUMNS_FIELD,
                "TST.TableModelArchive.number_of_header_columns",
            ),
            (
                FOOTER_ROWS_FIELD,
                "TST.TableModelArchive.number_of_footer_rows",
            ),
            (
                HEADER_ROWS_FROZEN_FIELD,
                "TST.TableModelArchive.header_rows_frozen",
            ),
            (
                HEADER_COLUMNS_FROZEN_FIELD,
                "TST.TableModelArchive.header_columns_frozen",
            ),
            (
                REPEATING_HEADER_ROWS_FIELD,
                "TST.TableModelArchive.repeating_header_rows_enabled",
            ),
            (
                REPEATING_HEADER_COLUMNS_FIELD,
                "TST.TableModelArchive.repeating_header_columns_enabled",
            ),
        ] {
            let mut source = dims.clone();
            source.extend(varint_field(field, 0));
            source.extend(varint_field(field, 0));
            assert_eq!(
                decode_table_header_settings(
                    &source,
                    DecodeOptions::new(source.len(), 16, source.len() * 2, 2)
                )
                .expect_err("duplicate")
                .duplicate_singular_field(),
                Some(name)
            );
        }
        for field in [TABLE_ROWS_FIELD, HEADER_ROWS_FIELD] {
            let mut source = dims.clone();
            push_varint((u64::from(field) << 3) | 5, &mut source);
            source.extend([0; 4]);
            assert!(
                decode_table_header_settings(
                    &source,
                    DecodeOptions::new(source.len(), 8, source.len() * 2, 2)
                )
                .is_err()
            );
        }
        for field in [HEADER_ROWS_FROZEN_FIELD, REPEATING_HEADER_COLUMNS_FIELD] {
            let mut source = dims.clone();
            push_varint((u64::from(field) << 3) | 2, &mut source);
            source.push(0);
            assert!(
                decode_table_header_settings(
                    &source,
                    DecodeOptions::new(source.len(), 8, source.len() * 2, 2)
                )
                .is_err()
            );
        }
        let mut overflow = dims.clone();
        overflow.extend(varint_field(HEADER_ROWS_FIELD, u64::from(u32::MAX) + 1));
        assert_eq!(
            decode_table_header_settings(
                &overflow,
                DecodeOptions::new(overflow.len(), 8, overflow.len() * 2, 2)
            )
            .expect_err("overflow")
            .noncanonical_reason(),
            Some("uint32 scalar exceeds u32")
        );
        let mut bool_two = dims.clone();
        bool_two.extend(varint_field(HEADER_ROWS_FROZEN_FIELD, 2));
        assert_eq!(
            decode_table_header_settings(
                &bool_two,
                DecodeOptions::new(bool_two.len(), 8, bool_two.len() * 2, 2)
            )
            .expect_err("bool")
            .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
        for source in [
            vec![0xb0, 0, 1, 0x38, 1],
            vec![0x30, 0x81, 0, 0x38, 1],
            vec![0x30],
        ] {
            assert!(
                decode_table_header_settings(
                    &source,
                    DecodeOptions::new(source.len(), 8, source.len() * 2, 2)
                )
                .is_err()
            );
        }
    }

    #[test]
    fn preserves_all_unknown_wire_forms() -> Result<(), DecodeError> {
        let mut source = dimensions();
        source.extend(varint_field(100, 1));
        push_varint((101_u64 << 3) | 1, &mut source);
        source.extend([0; 8]);
        push_varint((102_u64 << 3) | 2, &mut source);
        source.extend([2, 9, 8]);
        push_varint((103_u64 << 3) | 3, &mut source);
        source.extend([8, 1]);
        push_varint((103_u64 << 3) | 4, &mut source);
        push_varint((104_u64 << 3) | 5, &mut source);
        source.extend([0; 4]);
        assert!(
            decode_table_header_settings(
                &source,
                DecodeOptions::new(source.len(), 9, source.len() * 2, 2)
            )
            .is_ok()
        );
        let unclosed = [0x30, 1, 0x38, 1, 0xa3, 0x06];
        assert!(
            decode_table_header_settings(
                &unclosed,
                DecodeOptions::new(unclosed.len(), 3, unclosed.len() * 2, 2)
            )
            .is_err()
        );
        let noncanonical_length = [0x30, 1, 0x38, 1, 0xb2, 0x06, 0x80, 0, 0];
        assert_eq!(
            decode_table_header_settings(
                &noncanonical_length,
                DecodeOptions::new(
                    noncanonical_length.len(),
                    3,
                    noncanonical_length.len() * 2,
                    2
                )
            )
            .expect_err("length")
            .noncanonical_reason(),
            Some("length-delimited size")
        );
        Ok(())
    }
}
