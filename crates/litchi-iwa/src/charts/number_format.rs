//! Shared strict decimal-number formatting for native chart labels.
//!
//! Chart axes and series value labels store the same `TSK.FormatStructArchive`
//! payload in two legacy/current fields. This module owns the strongly typed
//! public values and the lossless dual-field codec used by both features.

use prost::Message;

use crate::protobuf::tsk;
use crate::wire::{
    append_varint_field, parse_wire_fields, patch_length_delimited_field, patch_varint_field,
};
use crate::{Error, Result};

const FORMAT_TYPE_FIELD: u32 = 1;
const DECIMAL_PLACES_FIELD: u32 = 2;
const NEGATIVE_STYLE_FIELD: u32 = 4;
const THOUSANDS_SEPARATOR_FIELD: u32 = 5;
const PREFIX_FIELD: u32 = 10_000;
const SUFFIX_FIELD: u32 = 10_001;

pub(crate) const NATIVE_NUMBER_FORMAT_TYPE: u64 = 2;
const NATIVE_DECIMAL_NUMBER_FORMAT: u64 = 256;
const NATIVE_AUTOMATIC_DECIMAL_PLACES: u64 = 253;
const NATIVE_MINUS_SIGN_NEGATIVE_STYLE: u64 = 0;
const NATIVE_PARENTHESES_NEGATIVE_STYLE: u64 = 2;
const NATIVE_MAXIMUM_DECIMAL_PLACES: u8 = 30;

/// A fixed number of decimal places accepted by the native iWork inspector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ChartFixedDecimalPlaces(u8);

impl ChartFixedDecimalPlaces {
    /// No fractional digits.
    pub const ZERO: Self = Self(0);

    /// The largest value accepted by Pages, Numbers, and Keynote.
    pub const MAXIMUM: Self = Self(NATIVE_MAXIMUM_DECIMAL_PLACES);

    /// Build a fixed decimal-place count accepted by iWork.
    pub fn new(value: u8) -> Result<Self> {
        if value > NATIVE_MAXIMUM_DECIMAL_PLACES {
            return Err(Error::InvalidFormat(format!(
                "chart decimal places must not exceed {NATIVE_MAXIMUM_DECIMAL_PLACES}"
            )));
        }
        Ok(Self(value))
    }

    /// Return the decimal-place count shown by iWork.
    pub const fn value(self) -> u8 {
        self.0
    }
}

impl TryFrom<u8> for ChartFixedDecimalPlaces {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
        Self::new(value)
    }
}

/// Automatic or fixed decimal places for native chart labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ChartDecimalPlaces {
    /// Let iWork derive the necessary number of fractional digits.
    #[default]
    Automatic,
    /// Always render exactly this many fractional digits.
    Fixed(ChartFixedDecimalPlaces),
}

impl ChartDecimalPlaces {
    /// Build a fixed decimal-place setting.
    pub fn fixed(value: u8) -> Result<Self> {
        ChartFixedDecimalPlaces::new(value).map(Self::Fixed)
    }
}

/// Native negative-number presentation for chart labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ChartNegativeStyle {
    /// Render a leading minus sign, for example `-100`.
    #[default]
    MinusSign,
    /// Render the magnitude in parentheses, for example `(100)`.
    Parentheses,
}

/// Decimal number formatting applied to native chart labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartNumberFormat {
    decimal_places: ChartDecimalPlaces,
    negative_style: ChartNegativeStyle,
    thousands_separator: bool,
}

impl ChartNumberFormat {
    /// Native default retained for existing series value-label callers.
    pub const NATIVE_DEFAULT: Self = Self::SERIES_VALUE_LABEL_NATIVE_DEFAULT;

    /// Native default for per-series value labels.
    pub const SERIES_VALUE_LABEL_NATIVE_DEFAULT: Self = Self::new(
        ChartDecimalPlaces::Automatic,
        ChartNegativeStyle::MinusSign,
        true,
    );

    /// Native default for value-axis labels.
    pub const AXIS_NATIVE_DEFAULT: Self = Self::new(
        ChartDecimalPlaces::Automatic,
        ChartNegativeStyle::MinusSign,
        false,
    );

    /// Construct a complete decimal-number format.
    pub const fn new(
        decimal_places: ChartDecimalPlaces,
        negative_style: ChartNegativeStyle,
        thousands_separator: bool,
    ) -> Self {
        Self {
            decimal_places,
            negative_style,
            thousands_separator,
        }
    }

    /// Return the automatic or fixed fractional-digit setting.
    pub const fn decimal_places(self) -> ChartDecimalPlaces {
        self.decimal_places
    }

    /// Return the negative-number presentation.
    pub const fn negative_style(self) -> ChartNegativeStyle {
        self.negative_style
    }

    /// Whether labels include locale-aware thousands separators.
    pub const fn thousands_separator(self) -> bool {
        self.thousands_separator
    }
}

impl Default for ChartNumberFormat {
    fn default() -> Self {
        Self::SERIES_VALUE_LABEL_NATIVE_DEFAULT
    }
}

/// Text placed immediately before and after native chart labels.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct ChartLabelAffixes {
    prefix: Box<str>,
    suffix: Box<str>,
}

impl ChartLabelAffixes {
    /// Construct chart-label affixes.
    pub fn new(prefix: impl Into<Box<str>>, suffix: impl Into<Box<str>>) -> Self {
        Self {
            prefix: prefix.into(),
            suffix: suffix.into(),
        }
    }

    /// Text placed before each label.
    pub fn prefix(&self) -> &str {
        &self.prefix
    }

    /// Text placed after each label.
    pub fn suffix(&self) -> &str {
        &self.suffix
    }

    /// Whether neither affix contains text.
    pub fn is_empty(&self) -> bool {
        self.prefix.is_empty() && self.suffix.is_empty()
    }
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct DualNumberFormatFields {
    pub(crate) legacy: u32,
    pub(crate) current: u32,
    pub(crate) format_type: u32,
}

pub(crate) fn read_dual_number_format(
    extension: Option<&[u8]>,
    fields: DualNumberFormatFields,
    default: ChartNumberFormat,
    context: &str,
) -> Result<ChartNumberFormat> {
    let Some(extension) = extension else {
        return Ok(default);
    };
    if let Some(format_type) = strict_optional_varint(extension, fields.format_type, context)?
        && format_type != NATIVE_NUMBER_FORMAT_TYPE
    {
        return Err(Error::InvalidFormat(format!(
            "{context} number-format type {format_type} is unsupported"
        )));
    }
    let legacy = strict_optional_length_delimited(extension, fields.legacy, context)?
        .map(|format| read_number_format(format, default, context))
        .transpose()?;
    let current = strict_optional_length_delimited(extension, fields.current, context)?
        .map(|format| read_number_format(format, default, context))
        .transpose()?;
    match (legacy, current) {
        (Some(legacy), Some(current)) if legacy != current => Err(Error::InvalidFormat(format!(
            "legacy and current {context} number formats disagree"
        ))),
        (_, Some(current)) => Ok(current),
        (Some(legacy), None) => Ok(legacy),
        (None, None) => Ok(default),
    }
}

pub(crate) fn read_dual_affixes(
    extension: Option<&[u8]>,
    fields: DualNumberFormatFields,
    context: &str,
) -> Result<ChartLabelAffixes> {
    let Some(extension) = extension else {
        return Ok(ChartLabelAffixes::default());
    };
    if let Some(format_type) = strict_optional_varint(extension, fields.format_type, context)?
        && format_type != NATIVE_NUMBER_FORMAT_TYPE
    {
        return Err(Error::InvalidFormat(format!(
            "{context} number-format type {format_type} is unsupported"
        )));
    }
    let legacy = strict_optional_length_delimited(extension, fields.legacy, context)?
        .map(|format| read_affixes(format, context))
        .transpose()?;
    let current = strict_optional_length_delimited(extension, fields.current, context)?
        .map(|format| read_affixes(format, context))
        .transpose()?;
    match (legacy, current) {
        (Some(legacy), Some(current)) if legacy != current => Err(Error::InvalidFormat(format!(
            "legacy and current {context} number formats disagree on label affixes"
        ))),
        (_, Some(current)) => Ok(current),
        (Some(legacy), None) => Ok(legacy),
        (None, None) => Ok(ChartLabelAffixes::default()),
    }
}

/// Patch both native formatter representations.
///
/// `None` means the existing extension already represents `expected` and can
/// be retained byte-for-byte.
pub(crate) fn patch_dual_number_format(
    extension: &[u8],
    fields: DualNumberFormatFields,
    expected: ChartNumberFormat,
    default: ChartNumberFormat,
    context: &str,
) -> Result<Option<Vec<u8>>> {
    let legacy = strict_optional_length_delimited(extension, fields.legacy, context)?;
    let current = strict_optional_length_delimited(extension, fields.current, context)?;
    if legacy.is_none() && current.is_none() && expected == default {
        return Ok(None);
    }
    let legacy_present = legacy.is_some();
    let current_present = current.is_some();
    let seed = current
        .or(legacy)
        .map_or_else(|| canonical_number_format(default), <[u8]>::to_vec);
    let legacy_format = patch_number_format(
        legacy.map_or_else(|| seed.clone(), <[u8]>::to_vec),
        expected,
        context,
    )?;
    let current_format = patch_number_format(
        current.map_or_else(|| seed, <[u8]>::to_vec),
        expected,
        context,
    )?;
    let extension = patch_length_delimited_field(
        extension,
        fields.legacy,
        legacy_present,
        Some(legacy_format.as_slice()),
    )?;
    let extension = patch_length_delimited_field(
        &extension,
        fields.current,
        current_present,
        Some(current_format.as_slice()),
    )?;
    let format_type_present =
        strict_optional_varint(&extension, fields.format_type, context)?.is_some();
    patch_varint_field(
        &extension,
        fields.format_type,
        format_type_present,
        Some(NATIVE_NUMBER_FORMAT_TYPE),
    )
    .map(Some)
}

/// Patch label affixes in both native formatter representations.
///
/// `None` means the existing extension already represents `expected` and can
/// be retained byte-for-byte.
pub(crate) fn patch_dual_affixes(
    extension: &[u8],
    fields: DualNumberFormatFields,
    expected: &ChartLabelAffixes,
    default_format: ChartNumberFormat,
    context: &str,
) -> Result<Option<Vec<u8>>> {
    let legacy = strict_optional_length_delimited(extension, fields.legacy, context)?;
    let current = strict_optional_length_delimited(extension, fields.current, context)?;
    let companion_type = strict_optional_varint(extension, fields.format_type, context)?;
    if let Some(format_type) = companion_type
        && format_type != NATIVE_NUMBER_FORMAT_TYPE
    {
        return Err(Error::InvalidFormat(format!(
            "{context} number-format type {format_type} is unsupported"
        )));
    }
    if legacy.is_none() && current.is_none() && expected.is_empty() {
        return Ok(None);
    }
    let legacy_present = legacy.is_some();
    let current_present = current.is_some();
    let seed = current
        .or(legacy)
        .map_or_else(|| canonical_number_format(default_format), <[u8]>::to_vec);
    let legacy_format = patch_affixes(
        legacy.map_or_else(|| seed.clone(), <[u8]>::to_vec),
        expected,
        context,
    )?;
    let current_format = patch_affixes(
        current.map_or_else(|| seed, <[u8]>::to_vec),
        expected,
        context,
    )?;
    if expected.is_empty()
        && legacy_present
        && current_present
        && companion_type == Some(NATIVE_NUMBER_FORMAT_TYPE)
    {
        let canonical_default = canonical_number_format(default_format);
        if legacy_format == canonical_default && current_format == canonical_default {
            return clear_dual_number_format(extension, fields, context);
        }
    }
    let extension = patch_length_delimited_field(
        extension,
        fields.legacy,
        legacy_present,
        Some(legacy_format.as_slice()),
    )?;
    let extension = patch_length_delimited_field(
        &extension,
        fields.current,
        current_present,
        Some(current_format.as_slice()),
    )?;
    if expected.is_empty() {
        return Ok(Some(extension));
    }
    let format_type_present = companion_type.is_some();
    patch_varint_field(
        &extension,
        fields.format_type,
        format_type_present,
        Some(NATIVE_NUMBER_FORMAT_TYPE),
    )
    .map(Some)
}

/// Remove both native formatter representations and the companion type.
///
/// `None` means none of the three fields was present.
pub(crate) fn clear_dual_number_format(
    extension: &[u8],
    fields: DualNumberFormatFields,
    context: &str,
) -> Result<Option<Vec<u8>>> {
    let legacy_present =
        strict_optional_length_delimited(extension, fields.legacy, context)?.is_some();
    let current_present =
        strict_optional_length_delimited(extension, fields.current, context)?.is_some();
    let format_type_present =
        strict_optional_varint(extension, fields.format_type, context)?.is_some();
    if !legacy_present && !current_present && !format_type_present {
        return Ok(None);
    }
    let extension = patch_length_delimited_field(extension, fields.legacy, legacy_present, None)?;
    let extension =
        patch_length_delimited_field(&extension, fields.current, current_present, None)?;
    patch_varint_field(&extension, fields.format_type, format_type_present, None).map(Some)
}

fn read_number_format(
    data: &[u8],
    default: ChartNumberFormat,
    context: &str,
) -> Result<ChartNumberFormat> {
    tsk::FormatStructArchive::decode(data)?;
    let format_type = strict_optional_varint(data, FORMAT_TYPE_FIELD, context)?
        .unwrap_or(NATIVE_DECIMAL_NUMBER_FORMAT);
    if format_type != NATIVE_DECIMAL_NUMBER_FORMAT {
        return Err(Error::InvalidFormat(format!(
            "{context} format type {format_type} is not a decimal number format"
        )));
    }
    let decimal_places = match strict_optional_varint(data, DECIMAL_PLACES_FIELD, context)?
        .unwrap_or(NATIVE_AUTOMATIC_DECIMAL_PLACES)
    {
        NATIVE_AUTOMATIC_DECIMAL_PLACES => ChartDecimalPlaces::Automatic,
        value if value <= u64::from(NATIVE_MAXIMUM_DECIMAL_PLACES) => ChartDecimalPlaces::Fixed(
            ChartFixedDecimalPlaces(u8::try_from(value).map_err(|_| {
                Error::InvalidFormat(format!(
                    "native {context} decimal places {value} cannot be represented"
                ))
            })?),
        ),
        value => {
            return Err(Error::InvalidFormat(format!(
                "native {context} decimal places {value} exceeds {NATIVE_MAXIMUM_DECIMAL_PLACES}"
            )));
        },
    };
    let negative_style = match strict_optional_varint(data, NEGATIVE_STYLE_FIELD, context)?
        .unwrap_or(NATIVE_MINUS_SIGN_NEGATIVE_STYLE)
    {
        NATIVE_MINUS_SIGN_NEGATIVE_STYLE => ChartNegativeStyle::MinusSign,
        NATIVE_PARENTHESES_NEGATIVE_STYLE => ChartNegativeStyle::Parentheses,
        value => {
            return Err(Error::InvalidFormat(format!(
                "unsupported native {context} negative style {value}"
            )));
        },
    };
    let thousands_separator =
        match strict_optional_varint(data, THOUSANDS_SEPARATOR_FIELD, context)?
            .unwrap_or(u64::from(default.thousands_separator()))
        {
            0 => false,
            1 => true,
            value => {
                return Err(Error::InvalidFormat(format!(
                    "native {context} thousands-separator flag {value} is not boolean"
                )));
            },
        };
    Ok(ChartNumberFormat::new(
        decimal_places,
        negative_style,
        thousands_separator,
    ))
}

fn read_affixes(data: &[u8], context: &str) -> Result<ChartLabelAffixes> {
    tsk::FormatStructArchive::decode(data)?;
    Ok(ChartLabelAffixes::new(
        strict_optional_string(data, PREFIX_FIELD, context)?.unwrap_or_default(),
        strict_optional_string(data, SUFFIX_FIELD, context)?.unwrap_or_default(),
    ))
}

fn patch_number_format(
    mut format: Vec<u8>,
    expected: ChartNumberFormat,
    context: &str,
) -> Result<Vec<u8>> {
    tsk::FormatStructArchive::decode(format.as_slice())?;
    format = patch_format_varint(
        &format,
        FORMAT_TYPE_FIELD,
        NATIVE_DECIMAL_NUMBER_FORMAT,
        context,
    )?;
    let decimal_places = match expected.decimal_places() {
        ChartDecimalPlaces::Automatic => NATIVE_AUTOMATIC_DECIMAL_PLACES,
        ChartDecimalPlaces::Fixed(value) => u64::from(value.value()),
    };
    format = patch_format_varint(&format, DECIMAL_PLACES_FIELD, decimal_places, context)?;
    let negative_style = match expected.negative_style() {
        ChartNegativeStyle::MinusSign => NATIVE_MINUS_SIGN_NEGATIVE_STYLE,
        ChartNegativeStyle::Parentheses => NATIVE_PARENTHESES_NEGATIVE_STYLE,
    };
    format = patch_format_varint(&format, NEGATIVE_STYLE_FIELD, negative_style, context)?;
    patch_format_varint(
        &format,
        THOUSANDS_SEPARATOR_FIELD,
        u64::from(expected.thousands_separator()),
        context,
    )
}

fn patch_affixes(
    mut format: Vec<u8>,
    expected: &ChartLabelAffixes,
    context: &str,
) -> Result<Vec<u8>> {
    tsk::FormatStructArchive::decode(format.as_slice())?;
    let prefix_present = strict_optional_string(&format, PREFIX_FIELD, context)?.is_some();
    format = patch_length_delimited_field(
        &format,
        PREFIX_FIELD,
        prefix_present,
        (!expected.prefix.is_empty()).then_some(expected.prefix.as_bytes()),
    )?;
    let suffix_present = strict_optional_string(&format, SUFFIX_FIELD, context)?.is_some();
    patch_length_delimited_field(
        &format,
        SUFFIX_FIELD,
        suffix_present,
        (!expected.suffix.is_empty()).then_some(expected.suffix.as_bytes()),
    )
}

fn patch_format_varint(
    data: &[u8],
    field_number: u32,
    value: u64,
    context: &str,
) -> Result<Vec<u8>> {
    let present = strict_optional_varint(data, field_number, context)?.is_some();
    patch_varint_field(data, field_number, present, Some(value))
}

fn canonical_number_format(default: ChartNumberFormat) -> Vec<u8> {
    let mut format = Vec::new();
    append_varint_field(&mut format, FORMAT_TYPE_FIELD, NATIVE_DECIMAL_NUMBER_FORMAT)
        .expect("writing to a Vec cannot fail");
    let decimal_places = match default.decimal_places() {
        ChartDecimalPlaces::Automatic => NATIVE_AUTOMATIC_DECIMAL_PLACES,
        ChartDecimalPlaces::Fixed(value) => u64::from(value.value()),
    };
    append_varint_field(&mut format, DECIMAL_PLACES_FIELD, decimal_places)
        .expect("writing to a Vec cannot fail");
    let negative_style = match default.negative_style() {
        ChartNegativeStyle::MinusSign => NATIVE_MINUS_SIGN_NEGATIVE_STYLE,
        ChartNegativeStyle::Parentheses => NATIVE_PARENTHESES_NEGATIVE_STYLE,
    };
    append_varint_field(&mut format, NEGATIVE_STYLE_FIELD, negative_style)
        .expect("writing to a Vec cannot fail");
    append_varint_field(
        &mut format,
        THOUSANDS_SEPARATOR_FIELD,
        u64::from(default.thousands_separator()),
    )
    .expect("writing to a Vec cannot fail");
    format
}

fn strict_optional_string(
    data: &[u8],
    field_number: u32,
    context: &str,
) -> Result<Option<Box<str>>> {
    strict_optional_length_delimited(data, field_number, context)?
        .map(|bytes| {
            std::str::from_utf8(bytes)
                .map(Box::<str>::from)
                .map_err(|error| {
                    Error::InvalidFormat(format!(
                        "{context} label-affix field {field_number} is not UTF-8: {error}"
                    ))
                })
        })
        .transpose()
}

fn strict_optional_length_delimited<'a>(
    data: &'a [u8],
    field_number: u32,
    context: &str,
) -> Result<Option<&'a [u8]>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular {context} number-format field {field_number} occurs more than once"
        )));
    }
    if field.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "{context} number-format field {field_number} is not length-delimited"
        )));
    }
    Ok(Some(&data[field.payload_start..field.end]))
}

fn strict_optional_varint(data: &[u8], field_number: u32, context: &str) -> Result<Option<u64>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular {context} number-format field {field_number} occurs more than once"
        )));
    }
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "{context} number-format field {field_number} is not a varint"
        )));
    }
    let (value, consumed) =
        litchi_iwa_common::varint::decode_varint_from_bytes(&data[field.key_end..field.end])
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "{context} number-format field {field_number} is invalid: {error}"
                ))
            })?;
    if consumed != field.end - field.key_end
        || litchi_iwa_common::varint::encoded_len(value) != consumed
    {
        return Err(Error::InvalidFormat(format!(
            "{context} number-format field {field_number} is not canonical"
        )));
    }
    Ok(Some(value))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::wire::{append_length_delimited_field, append_varint_field};

    const FIELDS: DualNumberFormatFields = DualNumberFormatFields {
        legacy: 2,
        current: 42,
        format_type: 3,
    };

    fn custom_format() -> ChartNumberFormat {
        ChartNumberFormat::new(
            ChartDecimalPlaces::fixed(2).unwrap(),
            ChartNegativeStyle::Parentheses,
            true,
        )
    }

    fn custom_affixes() -> ChartLabelAffixes {
        ChartLabelAffixes::new("USD ", " net")
    }

    #[test]
    fn fixed_decimal_places_match_the_native_inspector_range() {
        assert_eq!(
            ChartFixedDecimalPlaces::new(30).unwrap(),
            ChartFixedDecimalPlaces::MAXIMUM
        );
        assert!(ChartFixedDecimalPlaces::new(31).is_err());
    }

    #[test]
    fn dual_number_formats_round_trip_and_preserve_unknown_fields() {
        const UNKNOWN_FIELD: u32 = 9_001;
        let mut extension = Vec::new();
        append_varint_field(&mut extension, UNKNOWN_FIELD, 73).unwrap();
        let patched = patch_dual_number_format(
            &extension,
            FIELDS,
            custom_format(),
            ChartNumberFormat::AXIS_NATIVE_DEFAULT,
            "test axis",
        )
        .unwrap()
        .unwrap();
        assert_eq!(
            read_dual_number_format(
                Some(&patched),
                FIELDS,
                ChartNumberFormat::AXIS_NATIVE_DEFAULT,
                "test axis"
            )
            .unwrap(),
            custom_format()
        );
        assert_eq!(
            strict_optional_varint(&patched, UNKNOWN_FIELD, "test axis").unwrap(),
            Some(73)
        );
    }

    #[test]
    fn dual_affixes_round_trip_and_preserve_number_format_fields() {
        let formatted = patch_dual_number_format(
            &[],
            FIELDS,
            custom_format(),
            ChartNumberFormat::AXIS_NATIVE_DEFAULT,
            "test axis",
        )
        .unwrap()
        .unwrap();
        let patched = patch_dual_affixes(
            &formatted,
            FIELDS,
            &custom_affixes(),
            ChartNumberFormat::AXIS_NATIVE_DEFAULT,
            "test axis",
        )
        .unwrap()
        .unwrap();
        assert_eq!(
            read_dual_affixes(Some(&patched), FIELDS, "test axis").unwrap(),
            custom_affixes()
        );
        assert_eq!(
            read_dual_number_format(
                Some(&patched),
                FIELDS,
                ChartNumberFormat::AXIS_NATIVE_DEFAULT,
                "test axis"
            )
            .unwrap(),
            custom_format()
        );
        let cleared = patch_dual_affixes(
            &patched,
            FIELDS,
            &ChartLabelAffixes::default(),
            ChartNumberFormat::AXIS_NATIVE_DEFAULT,
            "test axis",
        )
        .unwrap()
        .unwrap();
        assert_eq!(
            read_dual_affixes(Some(&cleared), FIELDS, "test axis").unwrap(),
            ChartLabelAffixes::default()
        );
        assert_eq!(
            read_dual_number_format(
                Some(&cleared),
                FIELDS,
                ChartNumberFormat::AXIS_NATIVE_DEFAULT,
                "test axis"
            )
            .unwrap(),
            custom_format()
        );
    }

    #[test]
    fn dual_affixes_reject_conflicts_and_non_utf8_text() {
        let legacy = patch_affixes(
            canonical_number_format(ChartNumberFormat::AXIS_NATIVE_DEFAULT),
            &custom_affixes(),
            "test axis",
        )
        .unwrap();
        let current = canonical_number_format(ChartNumberFormat::AXIS_NATIVE_DEFAULT);
        let mut conflicting = Vec::new();
        append_length_delimited_field(&mut conflicting, FIELDS.legacy, &legacy).unwrap();
        append_length_delimited_field(&mut conflicting, FIELDS.current, &current).unwrap();
        assert!(read_dual_affixes(Some(&conflicting), FIELDS, "test axis").is_err());

        let mut malformed_format = canonical_number_format(ChartNumberFormat::AXIS_NATIVE_DEFAULT);
        append_length_delimited_field(&mut malformed_format, PREFIX_FIELD, &[0xff]).unwrap();
        let mut malformed = Vec::new();
        append_length_delimited_field(&mut malformed, FIELDS.current, &malformed_format).unwrap();
        assert!(read_dual_affixes(Some(&malformed), FIELDS, "test axis").is_err());
    }

    #[test]
    fn conflicting_and_malformed_dual_formats_are_rejected() {
        let legacy = canonical_number_format(ChartNumberFormat::AXIS_NATIVE_DEFAULT);
        let current = patch_number_format(
            canonical_number_format(ChartNumberFormat::AXIS_NATIVE_DEFAULT),
            custom_format(),
            "test axis",
        )
        .unwrap();
        let mut extension = Vec::new();
        append_length_delimited_field(&mut extension, FIELDS.legacy, &legacy).unwrap();
        append_length_delimited_field(&mut extension, FIELDS.current, &current).unwrap();
        assert!(
            read_dual_number_format(
                Some(&extension),
                FIELDS,
                ChartNumberFormat::AXIS_NATIVE_DEFAULT,
                "test axis"
            )
            .is_err()
        );

        let mut malformed = canonical_number_format(ChartNumberFormat::AXIS_NATIVE_DEFAULT);
        append_varint_field(&mut malformed, DECIMAL_PLACES_FIELD, 2).unwrap();
        assert!(
            read_number_format(
                &malformed,
                ChartNumberFormat::AXIS_NATIVE_DEFAULT,
                "test axis"
            )
            .is_err()
        );

        let mut wrong_type = Vec::new();
        append_varint_field(&mut wrong_type, FIELDS.format_type, 7).unwrap();
        assert!(
            read_dual_number_format(
                Some(&wrong_type),
                FIELDS,
                ChartNumberFormat::AXIS_NATIVE_DEFAULT,
                "test axis"
            )
            .is_err()
        );
    }
}
