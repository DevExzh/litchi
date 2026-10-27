//! The dual chart number-format codec (record 0375a): chart axes and series
//! value labels store the same `TSK.FormatStructArchive` in a legacy and a
//! current generated field, with a companion format-type field. Reads require
//! both copies to agree; writes keep them synchronised and keep every other
//! byte (the host's `charts::number_format`).

use litchi_iwa_common::chart::number_format::{
    DecimalPlaces, FixedDecimalPlaces, LabelAffixes, NegativeStyle, NumberFormat,
};
use litchi_iwa_common::wire;

use crate::graph::{self, Error, Result, push_varint_field};

const FORMAT_TYPE_FIELD: u32 = 1;
const DECIMAL_PLACES_FIELD: u32 = 2;
const NEGATIVE_STYLE_FIELD: u32 = 4;
const THOUSANDS_SEPARATOR_FIELD: u32 = 5;
const PREFIX_FIELD: u32 = 10_000;
const SUFFIX_FIELD: u32 = 10_001;
const NATIVE_NUMBER_FORMAT_TYPE: u64 = 2;
const NATIVE_DECIMAL_NUMBER_FORMAT: u64 = 256;
const NATIVE_AUTOMATIC_DECIMAL_PLACES: u64 = 253;
const NATIVE_MINUS_SIGN: u64 = 0;
const NATIVE_PARENTHESES: u64 = 2;
const NATIVE_MAXIMUM_DECIMAL_PLACES: u8 = 30;

/// The generated fields of one dual formatter.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Fields {
    /// The legacy `TSK.FormatStructArchive` field.
    pub legacy: u32,
    /// The current `TSK.FormatStructArchive` field.
    pub current: u32,
    /// The companion format-type field.
    pub format_type: u32,
}

impl Fields {
    /// The series value-label formatter of `TSCH.ChartSeriesNonStyleArchive`.
    pub const SERIES_VALUE_LABEL: Self = Self {
        legacy: 21,
        current: 98,
        format_type: 23,
    };
}

fn strict(data: &[u8], number: u32, wire_type: u8) -> Result<Option<&[u8]>> {
    let mut found = None;
    for field in graph::fields(data)? {
        if field.number() != number {
            continue;
        }
        if found.is_some() {
            return Err(Error::Invalid("a singular number-format field is repeated"));
        }
        if field.wire_type() != wire_type {
            return Err(Error::Invalid(
                "a number-format field has the wrong wire type",
            ));
        }
        let start = if wire_type == 2 {
            field.payload_start()
        } else {
            field.key_end()
        };
        found = Some(&data[start..field.end()]);
    }
    Ok(found)
}

fn strict_varint(data: &[u8], number: u32) -> Result<Option<u64>> {
    strict(data, number, 0)?
        .map(|bytes| {
            let (value, consumed) = litchi_iwa_common::decode_varint_from_bytes(bytes)
                .map_err(|_| Error::Invalid("a number-format varint is malformed"))?;
            if consumed != bytes.len() || litchi_iwa_common::varint::encoded_len(value) != consumed
            {
                return Err(Error::Invalid("a number-format varint is not canonical"));
            }
            Ok(value)
        })
        .transpose()
}

fn strict_string(data: &[u8], number: u32) -> Result<Option<&str>> {
    strict(data, number, 2)?
        .map(|bytes| {
            std::str::from_utf8(bytes).map_err(|_| Error::Invalid("a label affix is not UTF-8"))
        })
        .transpose()
}

fn check_type(extension: &[u8], fields: Fields) -> Result<Option<u64>> {
    let format_type = strict_varint(extension, fields.format_type)?;
    if format_type.is_some_and(|value| value != NATIVE_NUMBER_FORMAT_TYPE) {
        return Err(Error::Invalid("an unsupported number-format type"));
    }
    Ok(format_type)
}

fn read_format(data: &[u8], default: NumberFormat) -> Result<NumberFormat> {
    if strict_varint(data, FORMAT_TYPE_FIELD)?.unwrap_or(NATIVE_DECIMAL_NUMBER_FORMAT)
        != NATIVE_DECIMAL_NUMBER_FORMAT
    {
        return Err(Error::Invalid(
            "a label format is not a decimal number format",
        ));
    }
    let decimal_places = match strict_varint(data, DECIMAL_PLACES_FIELD)?
        .unwrap_or(NATIVE_AUTOMATIC_DECIMAL_PLACES)
    {
        NATIVE_AUTOMATIC_DECIMAL_PLACES => DecimalPlaces::Automatic,
        value if value <= u64::from(NATIVE_MAXIMUM_DECIMAL_PLACES) => DecimalPlaces::Fixed(
            FixedDecimalPlaces::new(value as u8)
                .map_err(|_| Error::Invalid("invalid decimal places"))?,
        ),
        _ => return Err(Error::Invalid("decimal places exceed the native maximum")),
    };
    let negative_style =
        match strict_varint(data, NEGATIVE_STYLE_FIELD)?.unwrap_or(NATIVE_MINUS_SIGN) {
            NATIVE_MINUS_SIGN => NegativeStyle::MinusSign,
            NATIVE_PARENTHESES => NegativeStyle::Parentheses,
            _ => return Err(Error::Invalid("an unsupported negative style")),
        };
    let thousands = match strict_varint(data, THOUSANDS_SEPARATOR_FIELD)?
        .unwrap_or(u64::from(default.thousands_separator()))
    {
        0 => false,
        1 => true,
        _ => return Err(Error::Invalid("a thousands-separator flag is not boolean")),
    };
    Ok(NumberFormat::new(decimal_places, negative_style, thousands))
}

fn read_affixes(data: &[u8]) -> Result<LabelAffixes> {
    LabelAffixes::new(
        strict_string(data, PREFIX_FIELD)?.unwrap_or_default(),
        strict_string(data, SUFFIX_FIELD)?.unwrap_or_default(),
    )
    .map_err(|_| Error::Invalid("invalid label affixes"))
}

fn dual<T: PartialEq>(
    extension: &[u8],
    fields: Fields,
    read: impl Fn(&[u8]) -> Result<T>,
) -> Result<Option<T>> {
    check_type(extension, fields)?;
    let legacy = strict(extension, fields.legacy, 2)?
        .map(&read)
        .transpose()?;
    let current = strict(extension, fields.current, 2)?
        .map(&read)
        .transpose()?;
    match (legacy, current) {
        (Some(legacy), Some(current)) if legacy != current => Err(Error::Invalid(
            "the legacy and current number formats disagree",
        )),
        (_, Some(current)) => Ok(Some(current)),
        (legacy, None) => Ok(legacy),
    }
}

/// The number format an extension stores (`default` when none).
///
/// # Errors
///
/// Malformed, unsupported or disagreeing formats.
pub fn number_format(
    extension: Option<&[u8]>,
    fields: Fields,
    default: NumberFormat,
) -> Result<NumberFormat> {
    let Some(extension) = extension else {
        return Ok(default);
    };
    Ok(dual(extension, fields, |data| read_format(data, default))?.unwrap_or(default))
}

/// The label affixes an extension stores (empty when none).
///
/// # Errors
///
/// Malformed, unsupported or disagreeing formats.
pub fn affixes(extension: Option<&[u8]>, fields: Fields) -> Result<LabelAffixes> {
    let Some(extension) = extension else {
        return Ok(LabelAffixes::default());
    };
    Ok(dual(extension, fields, read_affixes)?.unwrap_or_default())
}

fn native_decimal_places(value: DecimalPlaces) -> u64 {
    match value {
        DecimalPlaces::Fixed(value) => u64::from(value.value()),
        _ => NATIVE_AUTOMATIC_DECIMAL_PLACES,
    }
}

fn native_negative_style(value: NegativeStyle) -> u64 {
    match value {
        NegativeStyle::Parentheses => NATIVE_PARENTHESES,
        _ => NATIVE_MINUS_SIGN,
    }
}

fn canonical(default: NumberFormat) -> Vec<u8> {
    let mut format = Vec::new();
    push_varint_field(&mut format, FORMAT_TYPE_FIELD, NATIVE_DECIMAL_NUMBER_FORMAT);
    push_varint_field(
        &mut format,
        DECIMAL_PLACES_FIELD,
        native_decimal_places(default.decimal_places()),
    );
    push_varint_field(
        &mut format,
        NEGATIVE_STYLE_FIELD,
        native_negative_style(default.negative_style()),
    );
    push_varint_field(
        &mut format,
        THOUSANDS_SEPARATOR_FIELD,
        u64::from(default.thousands_separator()),
    );
    format
}

fn patch_varint(data: &[u8], number: u32, value: Option<u64>) -> Result<Vec<u8>> {
    let present = strict_varint(data, number)?.is_some();
    Ok(wire::patch_varint_field(data, number, present, value)?)
}

fn patch_format(format: &[u8], expected: NumberFormat) -> Result<Vec<u8>> {
    let format = patch_varint(
        format,
        FORMAT_TYPE_FIELD,
        Some(NATIVE_DECIMAL_NUMBER_FORMAT),
    )?;
    let format = patch_varint(
        &format,
        DECIMAL_PLACES_FIELD,
        Some(native_decimal_places(expected.decimal_places())),
    )?;
    let format = patch_varint(
        &format,
        NEGATIVE_STYLE_FIELD,
        Some(native_negative_style(expected.negative_style())),
    )?;
    patch_varint(
        &format,
        THOUSANDS_SEPARATOR_FIELD,
        Some(u64::from(expected.thousands_separator())),
    )
}

fn patch_affix_fields(format: &[u8], expected: &LabelAffixes) -> Result<Vec<u8>> {
    let present = strict_string(format, PREFIX_FIELD)?.is_some();
    let format = wire::patch_length_delimited_field(
        format,
        PREFIX_FIELD,
        present,
        (!expected.prefix().is_empty()).then_some(expected.prefix().as_bytes()),
    )?;
    let present = strict_string(&format, SUFFIX_FIELD)?.is_some();
    Ok(wire::patch_length_delimited_field(
        &format,
        SUFFIX_FIELD,
        present,
        (!expected.suffix().is_empty()).then_some(expected.suffix().as_bytes()),
    )?)
}

fn write_both(
    extension: &[u8],
    fields: Fields,
    default: NumberFormat,
    patch: impl Fn(&[u8]) -> Result<Vec<u8>>,
) -> Result<(Vec<u8>, Vec<u8>, bool, bool)> {
    let legacy = strict(extension, fields.legacy, 2)?;
    let current = strict(extension, fields.current, 2)?;
    let seed = current
        .or(legacy)
        .map_or_else(|| canonical(default), <[u8]>::to_vec);
    let legacy_format = patch(legacy.unwrap_or(&seed))?;
    let current_format = patch(current.unwrap_or(&seed))?;
    Ok((
        legacy_format,
        current_format,
        legacy.is_some(),
        current.is_some(),
    ))
}

fn store_both(
    extension: &[u8],
    fields: Fields,
    (legacy, current, legacy_present, current_present): (Vec<u8>, Vec<u8>, bool, bool),
) -> Result<Vec<u8>> {
    let extension = wire::patch_length_delimited_field(
        extension,
        fields.legacy,
        legacy_present,
        Some(&legacy),
    )?;
    Ok(wire::patch_length_delimited_field(
        &extension,
        fields.current,
        current_present,
        Some(&current),
    )?)
}

/// Write `expected` into both formatter copies (`None`: the extension
/// already stores it and is kept).
///
/// # Errors
///
/// Malformed or unsupported formats.
pub fn patch_number_format(
    extension: &[u8],
    fields: Fields,
    expected: NumberFormat,
    default: NumberFormat,
) -> Result<Option<Vec<u8>>> {
    if strict(extension, fields.legacy, 2)?.is_none()
        && strict(extension, fields.current, 2)?.is_none()
        && expected == default
    {
        return Ok(None);
    }
    let both = write_both(extension, fields, default, |format| {
        patch_format(format, expected)
    })?;
    let extension = store_both(extension, fields, both)?;
    Ok(Some(patch_varint(
        &extension,
        fields.format_type,
        Some(NATIVE_NUMBER_FORMAT_TYPE),
    )?))
}

/// Write `expected` affixes into both formatter copies (`None`: the
/// extension already stores them and is kept).
///
/// # Errors
///
/// Malformed or unsupported formats.
pub fn patch_affixes(
    extension: &[u8],
    fields: Fields,
    expected: &LabelAffixes,
    default: NumberFormat,
) -> Result<Option<Vec<u8>>> {
    let companion = check_type(extension, fields)?;
    let legacy_present = strict(extension, fields.legacy, 2)?.is_some();
    let current_present = strict(extension, fields.current, 2)?.is_some();
    if !legacy_present && !current_present && expected.is_empty() {
        return Ok(None);
    }
    let both = write_both(extension, fields, default, |format| {
        patch_affix_fields(format, expected)
    })?;
    if expected.is_empty()
        && legacy_present
        && current_present
        && companion == Some(NATIVE_NUMBER_FORMAT_TYPE)
    {
        let canonical_default = canonical(default);
        if both.0 == canonical_default && both.1 == canonical_default {
            return clear(extension, fields);
        }
    }
    let extension = store_both(extension, fields, both)?;
    if expected.is_empty() {
        return Ok(Some(extension));
    }
    Ok(Some(patch_varint(
        &extension,
        fields.format_type,
        Some(NATIVE_NUMBER_FORMAT_TYPE),
    )?))
}

/// Remove both formatter copies and the companion type (`None`: none was
/// stored).
///
/// # Errors
///
/// Malformed formats.
pub fn clear(extension: &[u8], fields: Fields) -> Result<Option<Vec<u8>>> {
    let legacy = strict(extension, fields.legacy, 2)?.is_some();
    let current = strict(extension, fields.current, 2)?.is_some();
    let format_type = strict_varint(extension, fields.format_type)?.is_some();
    if !legacy && !current && !format_type {
        return Ok(None);
    }
    let extension = wire::patch_length_delimited_field(extension, fields.legacy, legacy, None)?;
    let extension = wire::patch_length_delimited_field(&extension, fields.current, current, None)?;
    Ok(Some(patch_varint(&extension, fields.format_type, None)?))
}
