//! Lossless per-series value-label prefix and suffix CRUD for native charts.
//!
//! iWork persists both legacy and current number formatter objects for each
//! customized series. The two formatter representations must remain in sync.

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
const PREFIX_FIELD: u32 = 10_000;
const SUFFIX_FIELD: u32 = 10_001;

const NATIVE_NUMBER_FORMAT_TYPE: u64 = 2;
const NATIVE_DECIMAL_NUMBER_FORMAT: u64 = 256;
const NATIVE_AUTOMATIC_DECIMAL_PLACES: u64 = 253;
const NATIVE_MINUS_SIGN_NEGATIVE_STYLE: u64 = 0;

/// Text placed immediately before and after one chart series' value labels.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct ChartSeriesValueLabelAffixes {
    prefix: Box<str>,
    suffix: Box<str>,
}

impl ChartSeriesValueLabelAffixes {
    /// Construct value-label affixes.
    pub fn new(prefix: impl Into<Box<str>>, suffix: impl Into<Box<str>>) -> Self {
        Self {
            prefix: prefix.into(),
            suffix: suffix.into(),
        }
    }

    /// Text placed before each value label.
    pub fn prefix(&self) -> &str {
        &self.prefix
    }

    /// Text placed after each value label.
    pub fn suffix(&self) -> &str {
        &self.suffix
    }

    /// Whether neither affix contains text.
    pub fn is_empty(&self) -> bool {
        self.prefix.is_empty() && self.suffix.is_empty()
    }
}

/// Read value-label affixes in native series order.
pub(crate) fn chart_series_value_label_affixes(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
) -> Result<Vec<ChartSeriesValueLabelAffixes>> {
    ensure_supported_kind(kind)?;
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        ChartSeriesValueLabelAffixes::default(),
        read_affixes,
    )
}

/// Set value-label affixes in native series order.
pub(crate) fn set_chart_series_value_label_affixes(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    expected: &[ChartSeriesValueLabelAffixes],
) -> Result<()> {
    ensure_supported_kind(kind)?;
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "series value-label affixes",
        expected,
        ChartSeriesValueLabelAffixes::default(),
        read_affixes,
        patch_affixes,
    )
}

fn ensure_supported_kind(kind: ChartKind) -> Result<()> {
    if matches!(kind, ChartKind::Undefined | ChartKind::Unsupported(_)) {
        return Err(Error::InvalidFormat(format!(
            "chart kind {kind:?} has no supported series value-label affixes"
        )));
    }
    Ok(())
}

fn read_affixes(data: &[u8]) -> Result<ChartSeriesValueLabelAffixes> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(ChartSeriesValueLabelAffixes::default());
    };
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let legacy = strict_optional_length_delimited(extension, LEGACY_NUMBER_FORMAT_FIELD)?
        .map(read_format_affixes)
        .transpose()?;
    let current = strict_optional_length_delimited(extension, CURRENT_NUMBER_FORMAT_FIELD)?
        .map(read_format_affixes)
        .transpose()?;
    match (legacy, current) {
        (Some(legacy), Some(current)) if legacy != current => Err(Error::InvalidFormat(
            "legacy and current chart series number formats disagree on value-label affixes"
                .to_owned(),
        )),
        (_, Some(current)) => Ok(current),
        (Some(legacy), None) => Ok(legacy),
        (None, None) => Ok(ChartSeriesValueLabelAffixes::default()),
    }
}

fn read_format_affixes(data: &[u8]) -> Result<ChartSeriesValueLabelAffixes> {
    tsk::FormatStructArchive::decode(data)?;
    Ok(ChartSeriesValueLabelAffixes::new(
        strict_optional_string(data, PREFIX_FIELD)?.unwrap_or_default(),
        strict_optional_string(data, SUFFIX_FIELD)?.unwrap_or_default(),
    ))
}

fn patch_affixes(data: &[u8], expected: &ChartSeriesValueLabelAffixes) -> Result<Vec<u8>> {
    let existing_extension = generated_chart_series_non_style_extension(data)?;
    if existing_extension.is_none() && expected.is_empty() {
        return Ok(data.to_vec());
    }
    let mut extension = existing_extension.unwrap_or_default().to_vec();
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension.as_slice())?;

    let legacy = strict_optional_length_delimited(&extension, LEGACY_NUMBER_FORMAT_FIELD)?;
    let current = strict_optional_length_delimited(&extension, CURRENT_NUMBER_FORMAT_FIELD)?;
    let legacy_present = legacy.is_some();
    let current_present = current.is_some();
    if expected.is_empty() && legacy.is_none() && current.is_none() {
        return Ok(data.to_vec());
    }

    let seed = if let Some(format) = current.or(legacy) {
        format.to_vec()
    } else {
        canonical_number_format()?
    };
    let legacy_format = legacy.map_or_else(|| seed.clone(), <[u8]>::to_vec);
    let current_format = current.map_or_else(|| seed, <[u8]>::to_vec);
    let legacy_format = patch_format_affixes(legacy_format, expected)?;
    let current_format = patch_format_affixes(current_format, expected)?;
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
    if !expected.is_empty() {
        let format_type_present =
            strict_optional_varint(&extension, NUMBER_FORMAT_TYPE_FIELD)?.is_some();
        extension = patch_varint_field(
            &extension,
            NUMBER_FORMAT_TYPE_FIELD,
            format_type_present,
            Some(NATIVE_NUMBER_FORMAT_TYPE),
        )?;
    }
    let patched = patch_chart_series_non_style_extension(
        data,
        existing_extension.is_some(),
        Some(extension.as_slice()),
    )?;
    if read_affixes(&patched)? != *expected {
        return Err(Error::InvalidFormat(
            "chart series value-label affix wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

fn patch_format_affixes(
    mut format: Vec<u8>,
    expected: &ChartSeriesValueLabelAffixes,
) -> Result<Vec<u8>> {
    tsk::FormatStructArchive::decode(format.as_slice())?;
    let prefix_present = strict_optional_string(&format, PREFIX_FIELD)?.is_some();
    format = patch_length_delimited_field(
        &format,
        PREFIX_FIELD,
        prefix_present,
        (!expected.prefix.is_empty()).then_some(expected.prefix.as_bytes()),
    )?;
    let suffix_present = strict_optional_string(&format, SUFFIX_FIELD)?.is_some();
    patch_length_delimited_field(
        &format,
        SUFFIX_FIELD,
        suffix_present,
        (!expected.suffix.is_empty()).then_some(expected.suffix.as_bytes()),
    )
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

fn strict_optional_string(data: &[u8], field_number: u32) -> Result<Option<Box<str>>> {
    strict_optional_length_delimited(data, field_number)?
        .map(|bytes| {
            std::str::from_utf8(bytes)
                .map(Box::<str>::from)
                .map_err(|error| {
                    Error::InvalidFormat(format!(
                        "chart series value-label affix field {field_number} is not UTF-8: {error}"
                    ))
                })
        })
        .transpose()
}

fn strict_optional_length_delimited(data: &[u8], field_number: u32) -> Result<Option<&[u8]>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart series value-label field {field_number} occurs more than once"
        )));
    }
    if field.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart series value-label field {field_number} is not length-delimited"
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
            "singular chart series value-label field {field_number} occurs more than once"
        )));
    }
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart series value-label field {field_number} is not a varint"
        )));
    }
    let (value, consumed) = crate::varint::decode_varint_from_bytes(
        &data[field.key_end..field.end],
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "chart series value-label field {field_number} is invalid: {error}"
        ))
    })?;
    if consumed != field.end - field.key_end
        || crate::varint::encode_varint(value).len() != consumed
    {
        return Err(Error::InvalidFormat(format!(
            "chart series value-label field {field_number} is not canonical"
        )));
    }
    Ok(Some(value))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::series_non_style::canonical_empty_chart_series_non_style_data;
    use crate::wire::append_length_delimited_field;

    #[test]
    fn affixes_round_trip_through_both_native_number_formats() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let expected = ChartSeriesValueLabelAffixes::new("$", " USD");
        let patched = patch_affixes(&original, &expected).unwrap();
        assert_eq!(read_affixes(&patched).unwrap(), expected);

        let extension = generated_chart_series_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        let legacy = strict_optional_length_delimited(extension, LEGACY_NUMBER_FORMAT_FIELD)
            .unwrap()
            .unwrap();
        let current = strict_optional_length_delimited(extension, CURRENT_NUMBER_FORMAT_FIELD)
            .unwrap()
            .unwrap();
        assert_eq!(read_format_affixes(legacy).unwrap(), expected);
        assert_eq!(read_format_affixes(current).unwrap(), expected);
        assert_eq!(
            strict_optional_varint(extension, NUMBER_FORMAT_TYPE_FIELD).unwrap(),
            Some(NATIVE_NUMBER_FORMAT_TYPE)
        );
    }

    #[test]
    fn clearing_affixes_preserves_neighboring_formatter_bytes() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let customized =
            patch_affixes(&original, &ChartSeriesValueLabelAffixes::new("€", " net")).unwrap();
        let extension = generated_chart_series_non_style_extension(&customized)
            .unwrap()
            .unwrap();
        let current = strict_optional_length_delimited(extension, CURRENT_NUMBER_FORMAT_FIELD)
            .unwrap()
            .unwrap();
        let mut with_unknown = current.to_vec();
        append_varint_field(&mut with_unknown, 9_001, 73).unwrap();
        let extension = patch_length_delimited_field(
            extension,
            CURRENT_NUMBER_FORMAT_FIELD,
            true,
            Some(with_unknown.as_slice()),
        )
        .unwrap();
        let customized =
            patch_chart_series_non_style_extension(&customized, true, Some(extension.as_slice()))
                .unwrap();

        let cleared = patch_affixes(&customized, &ChartSeriesValueLabelAffixes::default()).unwrap();
        assert_eq!(
            read_affixes(&cleared).unwrap(),
            ChartSeriesValueLabelAffixes::default()
        );
        let extension = generated_chart_series_non_style_extension(&cleared)
            .unwrap()
            .unwrap();
        let current = strict_optional_length_delimited(extension, CURRENT_NUMBER_FORMAT_FIELD)
            .unwrap()
            .unwrap();
        assert_eq!(strict_optional_varint(current, 9_001).unwrap(), Some(73));
        assert_eq!(
            strict_optional_varint(extension, NUMBER_FORMAT_TYPE_FIELD).unwrap(),
            Some(NATIVE_NUMBER_FORMAT_TYPE)
        );
    }

    #[test]
    fn conflicting_native_number_formats_are_rejected() {
        let mut legacy = canonical_number_format().unwrap();
        append_length_delimited_field(&mut legacy, PREFIX_FIELD, b"$").unwrap();
        let mut current = canonical_number_format().unwrap();
        append_length_delimited_field(&mut current, PREFIX_FIELD, b"EUR ").unwrap();
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
        assert!(read_affixes(&data).is_err());
    }

    #[test]
    fn malformed_affix_wire_data_is_rejected() {
        let mut format = canonical_number_format().unwrap();
        append_varint_field(&mut format, PREFIX_FIELD, 1).unwrap();
        let mut extension = Vec::new();
        append_length_delimited_field(&mut extension, CURRENT_NUMBER_FORMAT_FIELD, &format)
            .unwrap();
        let data = patch_chart_series_non_style_extension(
            &canonical_empty_chart_series_non_style_data().unwrap(),
            false,
            Some(extension.as_slice()),
        )
        .unwrap();
        assert!(read_affixes(&data).is_err());
    }
}
