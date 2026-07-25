//! Lossless per-series value-label number-format CRUD for native charts.
//!
//! iWork stores a legacy and a current formatter for every customized chart
//! series. This module validates that both representations agree and patches
//! them together while preserving affixes and every unrelated wire field.

use prost::Message;

use crate::charts::ChartKind;
use crate::charts::series_non_style::{
    chart_series_non_style_values, generated_chart_series_non_style_extension,
    patch_chart_series_non_style_extension, set_chart_series_non_style_values,
};
use crate::protobuf::{tsch, tsk};
use crate::wire::{
    append_varint_field, parse_wire_fields, patch_length_delimited_field, patch_varint_field,
};
use crate::{Error, IWorkPackage, Result};

const LEGACY_NUMBER_FORMAT_FIELD: u32 = 21;
const NUMBER_FORMAT_TYPE_FIELD: u32 = 23;
const CURRENT_NUMBER_FORMAT_FIELD: u32 = 98;

const FORMAT_TYPE_FIELD: u32 = 1;
const DECIMAL_PLACES_FIELD: u32 = 2;
const NEGATIVE_STYLE_FIELD: u32 = 4;
const THOUSANDS_SEPARATOR_FIELD: u32 = 5;

const NATIVE_NUMBER_FORMAT_TYPE: u64 = 2;
const NATIVE_DECIMAL_NUMBER_FORMAT: u64 = 256;
const NATIVE_AUTOMATIC_DECIMAL_PLACES: u64 = 253;
const NATIVE_MINUS_SIGN_NEGATIVE_STYLE: u64 = 0;
const NATIVE_PARENTHESES_NEGATIVE_STYLE: u64 = 2;
const NATIVE_MAXIMUM_DECIMAL_PLACES: u8 = 30;

/// A fixed number of decimal places accepted by the native iWork inspector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ChartSeriesValueLabelFixedDecimalPlaces(u8);

impl ChartSeriesValueLabelFixedDecimalPlaces {
    /// No fractional digits.
    pub const ZERO: Self = Self(0);

    /// The largest value accepted by the Pages, Numbers, and Keynote inspector.
    pub const MAXIMUM: Self = Self(NATIVE_MAXIMUM_DECIMAL_PLACES);

    /// Build a fixed decimal-place count accepted by iWork.
    pub fn new(value: u8) -> Result<Self> {
        if value > NATIVE_MAXIMUM_DECIMAL_PLACES {
            return Err(Error::InvalidFormat(format!(
                "chart series value-label decimal places must not exceed {NATIVE_MAXIMUM_DECIMAL_PLACES}"
            )));
        }
        Ok(Self(value))
    }

    /// Return the decimal-place count shown by iWork.
    pub const fn value(self) -> u8 {
        self.0
    }
}

impl TryFrom<u8> for ChartSeriesValueLabelFixedDecimalPlaces {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
        Self::new(value)
    }
}

/// Automatic or fixed decimal places for a chart series' value labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ChartSeriesValueLabelDecimalPlaces {
    /// Let iWork derive the necessary number of fractional digits.
    #[default]
    Automatic,
    /// Always render exactly this many fractional digits.
    Fixed(ChartSeriesValueLabelFixedDecimalPlaces),
}

impl ChartSeriesValueLabelDecimalPlaces {
    /// Build a fixed decimal-place setting.
    pub fn fixed(value: u8) -> Result<Self> {
        ChartSeriesValueLabelFixedDecimalPlaces::new(value).map(Self::Fixed)
    }
}

/// Native negative-number presentation for chart series value labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ChartSeriesValueLabelNegativeStyle {
    /// Render a leading minus sign, for example `-100`.
    #[default]
    MinusSign,
    /// Render the magnitude in parentheses, for example `(100)`.
    Parentheses,
}

/// Number formatting applied to one chart series' value labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartSeriesValueLabelNumberFormat {
    decimal_places: ChartSeriesValueLabelDecimalPlaces,
    negative_style: ChartSeriesValueLabelNegativeStyle,
    thousands_separator: bool,
}

impl ChartSeriesValueLabelNumberFormat {
    /// The format used by newly inserted native iWork charts.
    pub const NATIVE_DEFAULT: Self = Self::new(
        ChartSeriesValueLabelDecimalPlaces::Automatic,
        ChartSeriesValueLabelNegativeStyle::MinusSign,
        true,
    );

    /// Construct a complete value-label number format.
    pub const fn new(
        decimal_places: ChartSeriesValueLabelDecimalPlaces,
        negative_style: ChartSeriesValueLabelNegativeStyle,
        thousands_separator: bool,
    ) -> Self {
        Self {
            decimal_places,
            negative_style,
            thousands_separator,
        }
    }

    /// Return the automatic or fixed fractional-digit setting.
    pub const fn decimal_places(self) -> ChartSeriesValueLabelDecimalPlaces {
        self.decimal_places
    }

    /// Return the negative-number presentation.
    pub const fn negative_style(self) -> ChartSeriesValueLabelNegativeStyle {
        self.negative_style
    }

    /// Whether value labels include locale-aware thousands separators.
    pub const fn thousands_separator(self) -> bool {
        self.thousands_separator
    }
}

impl Default for ChartSeriesValueLabelNumberFormat {
    fn default() -> Self {
        Self::NATIVE_DEFAULT
    }
}

/// Read number formats in native series order.
pub(crate) fn chart_series_value_label_number_formats(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
) -> Result<Vec<ChartSeriesValueLabelNumberFormat>> {
    ensure_supported_kind(kind)?;
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        ChartSeriesValueLabelNumberFormat::NATIVE_DEFAULT,
        read_number_format,
    )
}

/// Set number formats in native series order.
pub(crate) fn set_chart_series_value_label_number_formats(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    expected: &[ChartSeriesValueLabelNumberFormat],
) -> Result<()> {
    ensure_supported_kind(kind)?;
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "series value-label number formats",
        expected,
        ChartSeriesValueLabelNumberFormat::NATIVE_DEFAULT,
        read_number_format,
        patch_number_format,
    )
}

fn ensure_supported_kind(kind: ChartKind) -> Result<()> {
    if matches!(kind, ChartKind::Undefined | ChartKind::Unsupported(_)) {
        return Err(Error::InvalidFormat(format!(
            "chart kind {kind:?} has no supported series value-label number format"
        )));
    }
    Ok(())
}

fn read_number_format(data: &[u8]) -> Result<ChartSeriesValueLabelNumberFormat> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(ChartSeriesValueLabelNumberFormat::NATIVE_DEFAULT);
    };
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let legacy = strict_optional_length_delimited(extension, LEGACY_NUMBER_FORMAT_FIELD)?
        .map(read_format)
        .transpose()?;
    let current = strict_optional_length_delimited(extension, CURRENT_NUMBER_FORMAT_FIELD)?
        .map(read_format)
        .transpose()?;
    match (legacy, current) {
        (Some(legacy), Some(current)) if legacy != current => Err(Error::InvalidFormat(
            "legacy and current chart series number formats disagree".to_owned(),
        )),
        (_, Some(current)) => Ok(current),
        (Some(legacy), None) => Ok(legacy),
        (None, None) => Ok(ChartSeriesValueLabelNumberFormat::NATIVE_DEFAULT),
    }
}

fn read_format(data: &[u8]) -> Result<ChartSeriesValueLabelNumberFormat> {
    tsk::FormatStructArchive::decode(data)?;
    let format_type =
        strict_optional_varint(data, FORMAT_TYPE_FIELD)?.unwrap_or(NATIVE_DECIMAL_NUMBER_FORMAT);
    if format_type != NATIVE_DECIMAL_NUMBER_FORMAT {
        return Err(Error::InvalidFormat(format!(
            "chart series value-label format type {format_type} is not a decimal number format"
        )));
    }
    let decimal_places = match strict_optional_varint(data, DECIMAL_PLACES_FIELD)?
        .unwrap_or(NATIVE_AUTOMATIC_DECIMAL_PLACES)
    {
        NATIVE_AUTOMATIC_DECIMAL_PLACES => ChartSeriesValueLabelDecimalPlaces::Automatic,
        value if value <= u64::from(NATIVE_MAXIMUM_DECIMAL_PLACES) => {
            ChartSeriesValueLabelDecimalPlaces::Fixed(ChartSeriesValueLabelFixedDecimalPlaces(
                u8::try_from(value).map_err(|_| {
                    Error::InvalidFormat(format!(
                        "native chart series value-label decimal places {value} cannot be represented"
                    ))
                })?,
            ))
        },
        value => {
            return Err(Error::InvalidFormat(format!(
                "native chart series value-label decimal places {value} exceeds {NATIVE_MAXIMUM_DECIMAL_PLACES}"
            )));
        },
    };
    let negative_style = match strict_optional_varint(data, NEGATIVE_STYLE_FIELD)?
        .unwrap_or(NATIVE_MINUS_SIGN_NEGATIVE_STYLE)
    {
        NATIVE_MINUS_SIGN_NEGATIVE_STYLE => ChartSeriesValueLabelNegativeStyle::MinusSign,
        NATIVE_PARENTHESES_NEGATIVE_STYLE => ChartSeriesValueLabelNegativeStyle::Parentheses,
        value => {
            return Err(Error::InvalidFormat(format!(
                "unsupported native chart series value-label negative style {value}"
            )));
        },
    };
    let thousands_separator = match strict_optional_varint(data, THOUSANDS_SEPARATOR_FIELD)?
        .unwrap_or(1)
    {
        0 => false,
        1 => true,
        value => {
            return Err(Error::InvalidFormat(format!(
                "native chart series value-label thousands-separator flag {value} is not boolean"
            )));
        },
    };
    Ok(ChartSeriesValueLabelNumberFormat::new(
        decimal_places,
        negative_style,
        thousands_separator,
    ))
}

fn patch_number_format(
    data: &[u8],
    expected: &ChartSeriesValueLabelNumberFormat,
) -> Result<Vec<u8>> {
    let existing_extension = generated_chart_series_non_style_extension(data)?;
    if existing_extension.is_none()
        && *expected == ChartSeriesValueLabelNumberFormat::NATIVE_DEFAULT
    {
        return Ok(data.to_vec());
    }
    let mut extension = existing_extension.unwrap_or_default().to_vec();
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension.as_slice())?;
    let legacy = strict_optional_length_delimited(&extension, LEGACY_NUMBER_FORMAT_FIELD)?;
    let current = strict_optional_length_delimited(&extension, CURRENT_NUMBER_FORMAT_FIELD)?;
    if legacy.is_none()
        && current.is_none()
        && *expected == ChartSeriesValueLabelNumberFormat::NATIVE_DEFAULT
    {
        return Ok(data.to_vec());
    }
    let legacy_present = legacy.is_some();
    let current_present = current.is_some();
    let seed = current
        .or(legacy)
        .map_or_else(canonical_number_format, |format| Ok(format.to_vec()))?;
    let legacy_format = patch_format(
        legacy.map_or_else(|| seed.clone(), <[u8]>::to_vec),
        *expected,
    )?;
    let current_format = patch_format(current.map_or_else(|| seed, <[u8]>::to_vec), *expected)?;
    extension = patch_length_delimited_field(
        &extension,
        LEGACY_NUMBER_FORMAT_FIELD,
        legacy_present,
        Some(legacy_format.as_slice()),
    )?;
    extension = patch_length_delimited_field(
        &extension,
        CURRENT_NUMBER_FORMAT_FIELD,
        current_present,
        Some(current_format.as_slice()),
    )?;
    let format_type_present =
        strict_optional_varint(&extension, NUMBER_FORMAT_TYPE_FIELD)?.is_some();
    extension = patch_varint_field(
        &extension,
        NUMBER_FORMAT_TYPE_FIELD,
        format_type_present,
        Some(NATIVE_NUMBER_FORMAT_TYPE),
    )?;
    let patched = patch_chart_series_non_style_extension(
        data,
        existing_extension.is_some(),
        Some(extension.as_slice()),
    )?;
    if read_number_format(&patched)? != *expected {
        return Err(Error::InvalidFormat(
            "chart series value-label number-format wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

fn patch_format(
    mut format: Vec<u8>,
    expected: ChartSeriesValueLabelNumberFormat,
) -> Result<Vec<u8>> {
    tsk::FormatStructArchive::decode(format.as_slice())?;
    format = patch_format_varint(&format, FORMAT_TYPE_FIELD, NATIVE_DECIMAL_NUMBER_FORMAT)?;
    let decimal_places = match expected.decimal_places {
        ChartSeriesValueLabelDecimalPlaces::Automatic => NATIVE_AUTOMATIC_DECIMAL_PLACES,
        ChartSeriesValueLabelDecimalPlaces::Fixed(value) => u64::from(value.value()),
    };
    format = patch_format_varint(&format, DECIMAL_PLACES_FIELD, decimal_places)?;
    let negative_style = match expected.negative_style {
        ChartSeriesValueLabelNegativeStyle::MinusSign => NATIVE_MINUS_SIGN_NEGATIVE_STYLE,
        ChartSeriesValueLabelNegativeStyle::Parentheses => NATIVE_PARENTHESES_NEGATIVE_STYLE,
    };
    format = patch_format_varint(&format, NEGATIVE_STYLE_FIELD, negative_style)?;
    patch_format_varint(
        &format,
        THOUSANDS_SEPARATOR_FIELD,
        u64::from(expected.thousands_separator),
    )
}

fn patch_format_varint(data: &[u8], field_number: u32, value: u64) -> Result<Vec<u8>> {
    let present = strict_optional_varint(data, field_number)?.is_some();
    patch_varint_field(data, field_number, present, Some(value))
}

fn canonical_number_format() -> Result<Vec<u8>> {
    let mut format = Vec::new();
    append_varint_field(&mut format, FORMAT_TYPE_FIELD, NATIVE_DECIMAL_NUMBER_FORMAT)?;
    append_varint_field(
        &mut format,
        DECIMAL_PLACES_FIELD,
        NATIVE_AUTOMATIC_DECIMAL_PLACES,
    )?;
    append_varint_field(
        &mut format,
        NEGATIVE_STYLE_FIELD,
        NATIVE_MINUS_SIGN_NEGATIVE_STYLE,
    )?;
    append_varint_field(&mut format, THOUSANDS_SEPARATOR_FIELD, 1)?;
    Ok(format)
}

fn strict_optional_length_delimited(data: &[u8], field_number: u32) -> Result<Option<&[u8]>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart series value-label number-format field {field_number} occurs more than once"
        )));
    }
    if field.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart series value-label number-format field {field_number} is not length-delimited"
        )));
    }
    Ok(Some(&data[field.payload_start..field.end]))
}

fn strict_optional_varint(data: &[u8], field_number: u32) -> Result<Option<u64>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart series value-label number-format field {field_number} occurs more than once"
        )));
    }
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart series value-label number-format field {field_number} is not a varint"
        )));
    }
    let (value, consumed) = crate::varint::decode_varint_from_bytes(
        &data[field.key_end..field.end],
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "chart series value-label number-format field {field_number} is invalid: {error}"
        ))
    })?;
    if consumed != field.end - field.key_end
        || crate::varint::encode_varint(value).len() != consumed
    {
        return Err(Error::InvalidFormat(format!(
            "chart series value-label number-format field {field_number} is not canonical"
        )));
    }
    Ok(Some(value))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::series_non_style::canonical_empty_chart_series_non_style_data;
    use crate::wire::{append_length_delimited_field, append_varint_field};

    fn custom_format() -> ChartSeriesValueLabelNumberFormat {
        ChartSeriesValueLabelNumberFormat::new(
            ChartSeriesValueLabelDecimalPlaces::fixed(2).unwrap(),
            ChartSeriesValueLabelNegativeStyle::Parentheses,
            false,
        )
    }

    #[test]
    fn fixed_decimal_places_match_the_native_inspector_range() {
        assert_eq!(
            ChartSeriesValueLabelFixedDecimalPlaces::new(30).unwrap(),
            ChartSeriesValueLabelFixedDecimalPlaces::MAXIMUM
        );
        assert!(ChartSeriesValueLabelFixedDecimalPlaces::new(31).is_err());
    }

    #[test]
    fn number_format_round_trips_through_both_native_representations() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let expected = custom_format();
        let patched = patch_number_format(&original, &expected).unwrap();
        assert_eq!(read_number_format(&patched).unwrap(), expected);

        let extension = generated_chart_series_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        for field in [LEGACY_NUMBER_FORMAT_FIELD, CURRENT_NUMBER_FORMAT_FIELD] {
            let format = strict_optional_length_delimited(extension, field)
                .unwrap()
                .unwrap();
            assert_eq!(read_format(format).unwrap(), expected);
        }
        assert_eq!(
            strict_optional_varint(extension, NUMBER_FORMAT_TYPE_FIELD).unwrap(),
            Some(NATIVE_NUMBER_FORMAT_TYPE)
        );
    }

    #[test]
    fn number_format_patch_preserves_affixes_and_unknown_fields() {
        const PREFIX_FIELD: u32 = 10_000;
        const UNKNOWN_FIELD: u32 = 9_001;
        let mut format = canonical_number_format().unwrap();
        append_length_delimited_field(&mut format, PREFIX_FIELD, b"$").unwrap();
        append_varint_field(&mut format, UNKNOWN_FIELD, 73).unwrap();
        let mut extension = Vec::new();
        append_length_delimited_field(&mut extension, CURRENT_NUMBER_FORMAT_FIELD, &format)
            .unwrap();
        let original = patch_chart_series_non_style_extension(
            &canonical_empty_chart_series_non_style_data().unwrap(),
            false,
            Some(extension.as_slice()),
        )
        .unwrap();

        let patched = patch_number_format(&original, &custom_format()).unwrap();
        let extension = generated_chart_series_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        let current = strict_optional_length_delimited(extension, CURRENT_NUMBER_FORMAT_FIELD)
            .unwrap()
            .unwrap();
        assert_eq!(
            strict_optional_length_delimited(current, PREFIX_FIELD).unwrap(),
            Some(b"$".as_slice())
        );
        assert_eq!(
            strict_optional_varint(current, UNKNOWN_FIELD).unwrap(),
            Some(73)
        );
    }

    #[test]
    fn conflicting_and_malformed_native_formats_are_rejected() {
        let legacy = canonical_number_format().unwrap();
        let mut current = canonical_number_format().unwrap();
        current = patch_format_varint(
            &current,
            NEGATIVE_STYLE_FIELD,
            NATIVE_PARENTHESES_NEGATIVE_STYLE,
        )
        .unwrap();
        let mut extension = Vec::new();
        append_length_delimited_field(&mut extension, LEGACY_NUMBER_FORMAT_FIELD, &legacy).unwrap();
        append_length_delimited_field(&mut extension, CURRENT_NUMBER_FORMAT_FIELD, &current)
            .unwrap();
        let data = patch_chart_series_non_style_extension(
            &canonical_empty_chart_series_non_style_data().unwrap(),
            false,
            Some(extension.as_slice()),
        )
        .unwrap();
        assert!(read_number_format(&data).is_err());

        let mut malformed = canonical_number_format().unwrap();
        append_varint_field(&mut malformed, DECIMAL_PLACES_FIELD, 2).unwrap();
        assert!(read_format(&malformed).is_err());
    }
}
