//! Strict borrowed-name projections for Numbers sheet, form-sheet, and table models.

use std::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_numbers_names_generated::LitchiIwaProjection as projection;

const SHEET_NAME_FIELD: u32 = 1;
const FORM_SHEET_SUPER_FIELD: u32 = 1;
const TABLE_MODEL_ID_FIELD: u32 = 1;
const TABLE_MODEL_NAME_FIELD: u32 = 8;
const MAX_RECURSION: u32 = 64;

/// Finite resource policy for one Numbers name payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    bytes: usize,
    total_input_bytes: usize,
    fields: usize,
    work: usize,
    recursion: u32,
}
impl DecodeOptions {
    #[must_use]
    pub const fn new(bytes: usize, fields: usize, work: usize, recursion: u32) -> Self {
        Self {
            bytes,
            total_input_bytes: usize::MAX,
            fields,
            work,
            recursion,
        }
    }
    /// Bound the aggregate source bytes visited by one decode, including
    /// both the outer and nested messages of a form-based sheet.
    ///
    /// The default preserves the existing per-message byte policy. This
    /// additional ceiling lets callers share a residual budget across codecs.
    #[must_use]
    pub const fn with_total_input_bytes(mut self, maximum: usize) -> Self {
        self.total_input_bytes = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.bytes)
            .with_unknown_field_limit(self.fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion)
    }
}

/// A borrowed required Numbers sheet name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SheetNameSnapshot<'source> {
    name: &'source str,
}
impl<'source> SheetNameSnapshot<'source> {
    #[must_use]
    pub const fn name(self) -> &'source str {
        self.name
    }
}

/// Borrowed required identity and display name of a Numbers table model.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableNamesSnapshot<'source> {
    table_id: &'source str,
    table_name: &'source str,
}
impl<'source> TableNamesSnapshot<'source> {
    #[must_use]
    pub const fn table_id(self) -> &'source str {
        self.table_id
    }
    #[must_use]
    pub const fn table_name(self) -> &'source str {
        self.table_name
    }
}

/// Exact strict traversal cost for one successful Numbers name decode.
///
/// The byte total counts every source message passed through the strict
/// scanner.  Form-based sheets therefore include both the outer envelope and
/// its nested `super` payload.  The work total retains the codec's existing
/// two-pass accounting for the strict scanner and borrowed Buffa parity view;
/// the field total counts strict visits.  Buffa's borrowed parity projection
/// performs no owned allocation and is covered by the caller's existing wire
/// admission policy.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DecodeReport {
    input_bytes: usize,
    fields: usize,
    work: usize,
}

impl DecodeReport {
    /// Total source bytes visited by the strict scanner.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Number of strict wire fields visited, including unknown fields.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate strict-plus-Buffa work charged for the visited source
    /// messages.
    #[must_use]
    pub const fn work(self) -> usize {
        self.work
    }
}

/// Content-free byte or nesting resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    Bytes { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
}
/// Strict preflight or lazy-projection failure.
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
    const fn canonical(x: &'static str) -> Self {
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
            Kind::Resource(WireResourceLimit::Bytes { .. }) => {
                f.write_str("Numbers name byte limit exceeded")
            },
            Kind::Resource(WireResourceLimit::Nesting { .. }) => {
                f.write_str("Numbers name nesting limit exceeded")
            },
            Kind::Missing(x) => write!(f, "missing required field {x}"),
            Kind::Duplicate(x) => write!(f, "duplicate singular field {x}"),
            Kind::NonCanonical(x) => write!(f, "non-canonical protobuf representation: {x}"),
            Kind::Field { observed, maximum } => {
                write!(f, "visited {observed} fields; maximum is {maximum}")
            },
            Kind::Work { observed, maximum } => {
                write!(f, "requires {observed} work bytes; maximum is {maximum}")
            },
            Kind::Projection => {
                f.write_str("Numbers name strict preflight disagrees with Buffa projection")
            },
        }
    }
}
impl std::error::Error for DecodeError {}
impl From<buffa::DecodeError> for DecodeError {
    fn from(x: buffa::DecodeError) -> Self {
        match x {
            buffa::DecodeError::MessageTooLarge | buffa::DecodeError::RecursionLimitExceeded => {
                Self(Kind::Projection)
            },
            x => Self(Kind::Wire(x)),
        }
    }
}

/// Decode one `TN.SheetArchive` name without allocating an owned string.
pub fn decode_sheet_name(
    source: &[u8],
    options: DecodeOptions,
) -> Result<SheetNameSnapshot<'_>, DecodeError> {
    Ok(decode_sheet_name_with_report(source, options)?.0)
}

/// Decode one sheet name and return exact strict traversal accounting.
pub fn decode_sheet_name_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(SheetNameSnapshot<'_>, DecodeReport), DecodeError> {
    validate(source, options)?;
    let mut budget = Budget::new(options);
    let strict = sheet(source, options.recursion - 1, &mut budget)?;
    let view: projection::NumbersSheetArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    if view.name != strict.name {
        return Err(DecodeError(Kind::Projection));
    }
    Ok((strict, budget.report()))
}
/// Decode the nested `TN.FormBasedSheetArchive.super.name` without allocation.
pub fn decode_form_sheet_name(
    source: &[u8],
    options: DecodeOptions,
) -> Result<SheetNameSnapshot<'_>, DecodeError> {
    Ok(decode_form_sheet_name_with_report(source, options)?.0)
}

/// Decode a nested form-sheet name and return exact strict traversal
/// accounting.
pub fn decode_form_sheet_name_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(SheetNameSnapshot<'_>, DecodeReport), DecodeError> {
    validate(source, options)?;
    let mut budget = Budget::new(options);
    budget.message(source.len())?;
    let mut rest = source;
    let mut nested = None;
    while let Some(f) = next(&mut rest, options.recursion - 1, &mut budget)? {
        if f.number == FORM_SHEET_SUPER_FIELD {
            if nested.is_some() {
                return Err(DecodeError::duplicate("TN.FormBasedSheetArchive.super"));
            }
            nested = Some(f.bytes()?);
        }
    }
    let strict = sheet(
        nested.ok_or_else(|| DecodeError::missing("TN.FormBasedSheetArchive.super"))?,
        options
            .recursion
            .checked_sub(2)
            .ok_or_else(|| budget.nesting())?,
        &mut budget,
    )?;
    let view: projection::NumbersFormBasedSheetArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    let super_view = view.super_.get()?.ok_or(DecodeError(Kind::Projection))?;
    if super_view.name != strict.name {
        return Err(DecodeError(Kind::Projection));
    }
    Ok((strict, budget.report()))
}
/// Decode required `TST.TableModelArchive` identity and display names without allocation.
pub fn decode_table_names(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableNamesSnapshot<'_>, DecodeError> {
    Ok(decode_table_names_with_report(source, options)?.0)
}

/// Decode table identity/display names and return exact strict traversal
/// accounting.
pub fn decode_table_names_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableNamesSnapshot<'_>, DecodeReport), DecodeError> {
    validate(source, options)?;
    let mut budget = Budget::new(options);
    budget.message(source.len())?;
    let mut rest = source;
    let mut id = None;
    let mut name = None;
    while let Some(f) = next(&mut rest, options.recursion - 1, &mut budget)? {
        match f.number {
            TABLE_MODEL_ID_FIELD => {
                if id.is_some() {
                    return Err(DecodeError::duplicate("TST.TableModelArchive.table_id"));
                }
                id = Some(utf8(f.bytes()?)?);
            },
            TABLE_MODEL_NAME_FIELD => {
                if name.is_some() {
                    return Err(DecodeError::duplicate("TST.TableModelArchive.table_name"));
                }
                name = Some(utf8(f.bytes()?)?);
            },
            _ => {},
        }
    }
    let strict = TableNamesSnapshot {
        table_id: id.ok_or_else(|| DecodeError::missing("TST.TableModelArchive.table_id"))?,
        table_name: name.ok_or_else(|| DecodeError::missing("TST.TableModelArchive.table_name"))?,
    };
    let view: projection::NumbersTableModelArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    if view.table_id != strict.table_id || view.table_name != strict.table_name {
        return Err(DecodeError(Kind::Projection));
    }
    Ok((strict, budget.report()))
}

fn validate(source: &[u8], o: DecodeOptions) -> Result<(), DecodeError> {
    let hard =
        usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_| DecodeError(Kind::Projection))?;
    if o.bytes > hard {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: o.bytes,
            maximum: hard,
        }));
    }
    if source.len() > o.bytes {
        return Err(DecodeError::resource(WireResourceLimit::Bytes {
            observed: source.len(),
            maximum: o.bytes,
        }));
    }
    if o.recursion == 0 || o.recursion > MAX_RECURSION {
        return Err(DecodeError::resource(WireResourceLimit::Nesting {
            observed: o.recursion,
            maximum: MAX_RECURSION,
        }));
    }
    Ok(())
}
struct Budget {
    input_bytes: usize,
    max_input_bytes: usize,
    fields: usize,
    work: usize,
    max_fields: usize,
    max_work: usize,
    max_recursion: u32,
}
impl Budget {
    const fn new(o: DecodeOptions) -> Self {
        Self {
            input_bytes: 0,
            max_input_bytes: o.total_input_bytes,
            fields: 0,
            work: 0,
            max_fields: o.fields,
            max_work: o.work,
            max_recursion: o.recursion,
        }
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
    fn message(&mut self, n: usize) -> Result<(), DecodeError> {
        let input_bytes = self
            .input_bytes
            .checked_add(n)
            .ok_or(DecodeError(Kind::Projection))?;
        if input_bytes > self.max_input_bytes {
            return Err(DecodeError::resource(WireResourceLimit::Bytes {
                observed: input_bytes,
                maximum: self.max_input_bytes,
            }));
        }
        let observed = n
            .checked_mul(2)
            .and_then(|x| self.work.checked_add(x))
            .ok_or(DecodeError(Kind::Projection))?;
        if observed > self.max_work {
            return Err(DecodeError(Kind::Work {
                observed,
                maximum: self.max_work,
            }));
        }
        self.input_bytes = input_bytes;
        self.work = observed;
        Ok(())
    }
    const fn report(&self) -> DecodeReport {
        DecodeReport {
            input_bytes: self.input_bytes,
            fields: self.fields,
            work: self.work,
        }
    }
    const fn nesting(&self) -> DecodeError {
        DecodeError::resource(WireResourceLimit::Nesting {
            observed: self.max_recursion.saturating_add(1),
            maximum: self.max_recursion,
        })
    }
}
fn sheet<'a>(
    source: &'a [u8],
    depth: u32,
    b: &mut Budget,
) -> Result<SheetNameSnapshot<'a>, DecodeError> {
    b.message(source.len())?;
    let mut rest = source;
    let mut name = None;
    while let Some(f) = next(&mut rest, depth, b)? {
        if f.number == SHEET_NAME_FIELD {
            if name.is_some() {
                return Err(DecodeError::duplicate("TN.SheetArchive.name"));
            }
            name = Some(utf8(f.bytes()?)?);
        }
    }
    Ok(SheetNameSnapshot {
        name: name.ok_or_else(|| DecodeError::missing("TN.SheetArchive.name"))?,
    })
}
fn utf8(bytes: &[u8]) -> Result<&str, DecodeError> {
    str::from_utf8(bytes).map_err(|_| DecodeError::canonical("string is not valid UTF-8"))
}

#[derive(Clone, Copy)]
enum Value<'a> {
    Varint,
    Bytes(&'a [u8]),
    Other,
}
#[derive(Clone, Copy)]
struct Field<'a> {
    number: u32,
    wire: buffa::encoding::WireType,
    value: Value<'a>,
}
impl<'a> Field<'a> {
    fn bytes(self) -> Result<&'a [u8], DecodeError> {
        if self.wire != buffa::encoding::WireType::LengthDelimited {
            return Err(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected: 2,
                actual: self.wire as u8,
            }
            .into());
        }
        if let Value::Bytes(x) = self.value {
            Ok(x)
        } else {
            Err(DecodeError(Kind::Projection))
        }
    }
}
fn next<'a>(
    s: &mut &'a [u8],
    depth: u32,
    b: &mut Budget,
) -> Result<Option<Field<'a>>, DecodeError> {
    if s.is_empty() {
        return Ok(None);
    };
    let (tag, c) = varint(s)?;
    if !c {
        return Err(DecodeError::canonical("protobuf field key"));
    };
    b.field()?;
    let raw = u32::try_from(tag).map_err(|_| buffa::DecodeError::InvalidFieldNumber)?;
    let number = raw >> 3;
    if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    };
    let wire = buffa::encoding::WireType::from_u32(raw & 7)?;
    let value = match wire {
        buffa::encoding::WireType::Varint => {
            let (_, c) = varint(s)?;
            if !c {
                return Err(DecodeError::canonical("protobuf varint value"));
            };
            Value::Varint
        },
        buffa::encoding::WireType::Fixed32 => {
            take(s, 4)?;
            Value::Other
        },
        buffa::encoding::WireType::Fixed64 => {
            take(s, 8)?;
            Value::Other
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (n, c) = varint(s)?;
            if !c {
                return Err(DecodeError::canonical("length-delimited size"));
            };
            Value::Bytes(take(
                s,
                usize::try_from(n).map_err(|_| buffa::DecodeError::MessageTooLarge)?,
            )?)
        },
        buffa::encoding::WireType::StartGroup => {
            skip_group(
                s,
                number,
                depth.checked_sub(1).ok_or_else(|| b.nesting())?,
                b,
            )?;
            Value::Other
        },
        buffa::encoding::WireType::EndGroup => {
            return Err(buffa::DecodeError::InvalidEndGroup(number).into());
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw & 7).into()),
    };
    Ok(Some(Field {
        number,
        wire,
        value,
    }))
}
fn skip_group(s: &mut &[u8], number: u32, depth: u32, b: &mut Budget) -> Result<(), DecodeError> {
    loop {
        if s.is_empty() {
            return Err(buffa::DecodeError::UnexpectedEof.into());
        };
        let before = *s;
        let (tag, c) = varint(s)?;
        if !c {
            return Err(DecodeError::canonical("protobuf field key"));
        };
        let raw = u32::try_from(tag).map_err(|_| buffa::DecodeError::InvalidFieldNumber)?;
        if raw >> 3 == 0 {
            return Err(buffa::DecodeError::InvalidFieldNumber.into());
        };
        if raw & 7 == 4 {
            b.field()?;
            if raw >> 3 == number {
                return Ok(());
            }
            return Err(buffa::DecodeError::InvalidEndGroup(raw >> 3).into());
        };
        *s = before;
        let _ = next(s, depth, b)?;
    }
}
fn varint(s: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *s;
    let mut value = 0;
    for i in 0..10 {
        let byte = *original.get(i).ok_or(buffa::DecodeError::UnexpectedEof)?;
        if i == 9 && byte > 1 {
            return Err(buffa::DecodeError::VarintTooLong.into());
        };
        value |= u64::from(byte & 127) << (i * 7);
        if byte & 128 == 0 {
            *s = &original[i + 1..];
            let mut n = value;
            let mut len = 1;
            while n >= 128 {
                n >>= 7;
                len += 1
            }
            return Ok((value, len == i + 1));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}
fn take<'a>(s: &mut &'a [u8], n: usize) -> Result<&'a [u8], DecodeError> {
    if s.len() < n {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    };
    let (a, z) = s.split_at(n);
    *s = z;
    Ok(a)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(source.len(), 32, source.len() * 4, 4)
    }

    #[test]
    fn decodes_borrowed_sheet_form_and_table_names() -> Result<(), DecodeError> {
        let sheet = [0x0a, 0x05, b'S', b'h', b'e', b'e', b't'];
        let form = [0x0a, 0x07, 0x0a, 0x05, b'S', b'h', b'e', b'e', b't'];
        let table = [
            0x0a, 0x02, b'i', b'd', 0x42, 0x05, b'T', b'a', b'b', b'l', b'e',
        ];
        assert_eq!(decode_sheet_name(&sheet, options(&sheet))?.name(), "Sheet");
        assert_eq!(
            decode_form_sheet_name(&form, options(&form))?.name(),
            "Sheet"
        );
        let names = decode_table_names(&table, options(&table))?;
        assert_eq!((names.table_id(), names.table_name()), ("id", "Table"));
        Ok(())
    }

    #[test]
    fn reports_every_strict_name_source_pass() -> Result<(), DecodeError> {
        let sheet = [0x0a, 0x05, b'S', b'h', b'e', b'e', b't'];
        let form = [0x0a, 0x07, 0x0a, 0x05, b'S', b'h', b'e', b'e', b't'];
        let table = [
            0x0a, 0x02, b'i', b'd', 0x42, 0x05, b'T', b'a', b'b', b'l', b'e',
        ];

        let (_, report) = decode_sheet_name_with_report(&sheet, options(&sheet))?;
        assert_eq!(report.input_bytes(), sheet.len());
        assert_eq!(report.fields(), 1);
        assert_eq!(report.work(), sheet.len() * 2);

        let (_, report) = decode_form_sheet_name_with_report(&form, options(&form))?;
        assert_eq!(report.input_bytes(), form.len() + 7);
        assert_eq!(report.fields(), 2);
        assert_eq!(report.work(), (form.len() + 7) * 2);

        let (_, report) = decode_table_names_with_report(&table, options(&table))?;
        assert_eq!(report.input_bytes(), table.len());
        assert_eq!(report.fields(), 2);
        assert_eq!(report.work(), table.len() * 2);
        Ok(())
    }

    #[test]
    fn form_name_aggregate_byte_limit_precedes_nested_decoding() -> Result<(), DecodeError> {
        let form = [0x0a, 0x03, 0x0a, 0x01, b'N'];
        let total = form.len() + 3;
        let (_, report) = decode_form_sheet_name_with_report(
            &form,
            options(&form).with_total_input_bytes(total),
        )?;
        assert_eq!(report.input_bytes(), total);
        let error = decode_form_sheet_name_with_report(
            &form,
            options(&form).with_total_input_bytes(total - 1),
        )
        .unwrap_err();
        assert_eq!(
            error.wire_resource_limit(),
            Some(WireResourceLimit::Bytes {
                observed: total,
                maximum: total - 1,
            })
        );

        // The envelope is valid but its nested name has a truncated length.
        // Exhaustion must be detected before that nested payload is decoded.
        let malformed_nested = [0x0a, 0x02, 0x0a, 0x80];
        let error = decode_form_sheet_name_with_report(
            &malformed_nested,
            options(&malformed_nested).with_total_input_bytes(malformed_nested.len()),
        )
        .unwrap_err();
        assert_eq!(
            error.wire_resource_limit(),
            Some(WireResourceLimit::Bytes {
                observed: malformed_nested.len() + 2,
                maximum: malformed_nested.len(),
            })
        );
        Ok(())
    }

    #[test]
    fn strict_failures_and_exact_limits() {
        let duplicate = [0x0a, 1, b'a', 0x0a, 1, b'b'];
        assert_eq!(
            decode_sheet_name(&duplicate, options(&duplicate))
                .expect_err("duplicate")
                .duplicate_singular_field(),
            Some("TN.SheetArchive.name")
        );
        let invalid_utf8 = [0x0a, 1, 0xff];
        assert_eq!(
            decode_sheet_name(&invalid_utf8, options(&invalid_utf8))
                .expect_err("utf8")
                .noncanonical_reason(),
            Some("string is not valid UTF-8")
        );
        let wrong_wire = [0x08, 1];
        assert!(decode_sheet_name(&wrong_wire, options(&wrong_wire)).is_err());
        let sheet = [0x0a, 1, b'a'];
        assert!(decode_sheet_name(&sheet, DecodeOptions::new(3, 1, 6, 1)).is_ok());
        assert_eq!(
            decode_sheet_name(&sheet, DecodeOptions::new(2, 1, 6, 1))
                .expect_err("bytes")
                .wire_resource_limit(),
            Some(WireResourceLimit::Bytes {
                observed: 3,
                maximum: 2
            })
        );
        assert_eq!(
            decode_sheet_name(&sheet, DecodeOptions::new(3, 0, 6, 1))
                .expect_err("fields")
                .field_limit_values(),
            Some((1, 0))
        );
        assert_eq!(
            decode_sheet_name(&sheet, DecodeOptions::new(3, 1, 5, 1))
                .expect_err("work")
                .work_limit_values(),
            Some((6, 5))
        );
    }
}
