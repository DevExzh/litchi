//! Lossless per-series error-bar CRUD for native charts.
//!
//! This module is the IWA boundary only: it decodes and patches the native
//! protobuf extension, while archive-free semantic values live in
//! `litchi_iwa_common::chart::error_bar`.

use litchi_iwa_common::chart::error_bar::{
    CustomValue, CustomValues, Direction, FixedValue, Percentage, Series, StandardDeviationCount,
};
use prost::Message;

use crate::charts::series_non_style::{
    NewChartSeriesNonStyleBase, chart_series_non_style_values,
    generated_chart_series_non_style_extension, patch_chart_series_non_style_extension,
    set_chart_series_non_style_values,
};
use crate::protobuf::tsch;
use crate::wire::{
    parse_wire_fields, patch_fixed32_field, patch_length_delimited_field, patch_varint_field,
    repeated_fixed64_values, rewrite_repeated_fixed64_fields,
};
use crate::{Error, IWorkPackage, Result};

const CUSTOM_NEGATIVE_VALUES_FIELD: u32 = 2;
const CUSTOM_POSITIVE_VALUES_FIELD: u32 = 4;
const FIXED_VALUE_FIELD: u32 = 6;
const PERCENTAGE_FIELD: u32 = 8;
const DIRECTION_FIELD: u32 = 10;
const STANDARD_DEVIATION_FIELD: u32 = 12;
const VALUE_TYPE_FIELD: u32 = 14;
const SHOW_ERROR_BARS_FIELD: u32 = 27;
const CUSTOM_ARRAY_VALUE_FIELD: u32 = 1;

const NATIVE_DIRECTION_BOTH: i32 = 1;
const NATIVE_FIXED_VALUE: i32 = 1;
const NATIVE_PERCENTAGE: i32 = 2;
const NATIVE_STANDARD_DEVIATION: i32 = 3;
const NATIVE_STANDARD_ERROR: i32 = 4;
const NATIVE_CUSTOM_VALUES: i32 = 5;

pub(crate) fn chart_series_error_bars(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<Vec<Series>> {
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        Series::None,
        read_error_bars,
    )
}

pub(crate) fn set_chart_series_error_bars(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    expected: &[Series],
) -> Result<()> {
    for value in expected {
        value.validate()?;
    }
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "series error bars",
        NewChartSeriesNonStyleBase::Styled,
        expected,
        Series::None,
        read_error_bars,
        patch_error_bars,
    )
}

fn read_error_bars(data: &[u8]) -> Result<Series> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(Series::None);
    };
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let visible = strict_optional_bool(extension, SHOW_ERROR_BARS_FIELD)?.unwrap_or(false);
    let direction =
        strict_optional_i32(extension, DIRECTION_FIELD)?.unwrap_or(NATIVE_DIRECTION_BOTH);
    let native_type =
        strict_optional_i32(extension, VALUE_TYPE_FIELD)?.unwrap_or(NATIVE_FIXED_VALUE);
    let fixed = strict_optional_f32(extension, FIXED_VALUE_FIELD)?;
    let percentage = strict_optional_f32(extension, PERCENTAGE_FIELD)?;
    let deviations = strict_optional_f32(extension, STANDARD_DEVIATION_FIELD)?;
    let negative = read_custom_values(extension, CUSTOM_NEGATIVE_VALUES_FIELD)?;
    let positive = read_custom_values(extension, CUSTOM_POSITIVE_VALUES_FIELD)?;
    if !visible {
        return Ok(Series::None);
    }

    let direction = Direction::from_native(direction);
    let value = match native_type {
        NATIVE_FIXED_VALUE => Series::FixedValue {
            direction,
            value: FixedValue::new(fixed.unwrap_or(FixedValue::DEFAULT.value()))?,
        },
        NATIVE_PERCENTAGE => Series::Percentage {
            direction,
            percentage: Percentage::new(f32_to_u8(
                percentage.unwrap_or(f32::from(Percentage::DEFAULT.value())),
                "percentage",
            )?)?,
        },
        NATIVE_STANDARD_DEVIATION => Series::StandardDeviation {
            direction,
            deviations: StandardDeviationCount::new(f32_to_u32(
                deviations.unwrap_or(StandardDeviationCount::DEFAULT.value() as f32),
                "standard-deviation count",
            )?)?,
        },
        NATIVE_STANDARD_ERROR => Series::StandardError { direction },
        NATIVE_CUSTOM_VALUES => Series::CustomValues {
            direction,
            values: CustomValues::from_validated(positive, negative)?,
        },
        native_type => Series::unsupported(direction, native_type)?,
    };
    value.validate()?;
    Ok(value)
}

fn patch_error_bars(data: &[u8], expected: &Series) -> Result<Vec<u8>> {
    expected.validate()?;
    let existing_extension = generated_chart_series_non_style_extension(data)?;
    if existing_extension.is_none() && expected == &Series::None {
        return Ok(data.to_vec());
    }
    let mut extension = existing_extension.unwrap_or_default().to_vec();
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension.as_slice())?;
    let visible = expected.direction().is_some();
    extension = patch_optional_varint(&extension, SHOW_ERROR_BARS_FIELD, visible.then_some(1))?;
    if let Some(direction) = expected.direction() {
        extension = patch_optional_varint(
            &extension,
            DIRECTION_FIELD,
            Some(encode_i32(direction.native_value())),
        )?;
        extension = patch_optional_varint(
            &extension,
            VALUE_TYPE_FIELD,
            expected.native_type().map(encode_i32),
        )?;
        match expected {
            Series::FixedValue { value, .. } => {
                extension = patch_optional_fixed32(
                    &extension,
                    FIXED_VALUE_FIELD,
                    Some(value.value().to_bits()),
                )?;
            },
            Series::Percentage { percentage, .. } => {
                extension = patch_optional_fixed32(
                    &extension,
                    PERCENTAGE_FIELD,
                    Some(f32::from(percentage.value()).to_bits()),
                )?;
            },
            Series::StandardDeviation { deviations, .. } => {
                extension = patch_optional_fixed32(
                    &extension,
                    STANDARD_DEVIATION_FIELD,
                    Some((deviations.value() as f32).to_bits()),
                )?;
            },
            Series::CustomValues { values, .. } => {
                extension = patch_custom_values(
                    &extension,
                    CUSTOM_NEGATIVE_VALUES_FIELD,
                    values.negative(),
                )?;
                extension = patch_custom_values(
                    &extension,
                    CUSTOM_POSITIVE_VALUES_FIELD,
                    values.positive(),
                )?;
            },
            Series::None | Series::StandardError { .. } | Series::Unsupported { .. } => {},
            _ => {},
        }
    }

    let patched = patch_chart_series_non_style_extension(
        data,
        existing_extension.is_some(),
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    if read_error_bars(&patched)? != *expected {
        return Err(Error::InvalidFormat(
            "chart series error-bar wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

fn patch_custom_values(
    extension: &[u8],
    field_number: u32,
    expected: &[CustomValue],
) -> Result<Vec<u8>> {
    let existing = strict_optional_bytes(extension, field_number)?;
    let mut nested = existing.unwrap_or_default().to_vec();
    if !nested.is_empty() {
        tsch::ChartsNsArrayOfNsNumberDoubleArchive::decode(nested.as_slice())?;
    }
    nested = rewrite_repeated_fixed64_fields(
        &nested,
        CUSTOM_ARRAY_VALUE_FIELD,
        &expected
            .iter()
            .map(|value| value.value().to_bits())
            .collect::<Vec<_>>(),
    )?;
    patch_length_delimited_field(
        extension,
        field_number,
        existing.is_some(),
        (!nested.is_empty()).then_some(nested.as_slice()),
    )
}

fn read_custom_values(extension: &[u8], field_number: u32) -> Result<Vec<CustomValue>> {
    let Some(nested) = strict_optional_bytes(extension, field_number)? else {
        return Ok(Vec::new());
    };
    tsch::ChartsNsArrayOfNsNumberDoubleArchive::decode(nested)?;
    repeated_fixed64_values(nested, CUSTOM_ARRAY_VALUE_FIELD)?
        .into_iter()
        .map(|bits| CustomValue::new(f64::from_bits(bits)).map_err(Into::into))
        .collect()
}

fn f32_to_u8(value: f32, label: &str) -> Result<u8> {
    if !value.is_finite() || value.fract() != 0.0 || value < 0.0 || value > f32::from(u8::MAX) {
        return Err(Error::InvalidFormat(format!(
            "chart error-bar {label} is not an unsigned integer: {value}"
        )));
    }
    Ok(value as u8)
}

fn f32_to_u32(value: f32, label: &str) -> Result<u32> {
    if !value.is_finite() || value.fract() != 0.0 || value < 0.0 || value > u32::MAX as f32 {
        return Err(Error::InvalidFormat(format!(
            "chart error-bar {label} is not an unsigned integer: {value}"
        )));
    }
    Ok(value as u32)
}

fn patch_optional_varint(data: &[u8], field_number: u32, value: Option<u64>) -> Result<Vec<u8>> {
    let present = strict_optional_varint(data, field_number)?.is_some();
    patch_varint_field(data, field_number, present, value)
}

fn patch_optional_fixed32(data: &[u8], field_number: u32, value: Option<u32>) -> Result<Vec<u8>> {
    let present = strict_optional_fixed32(data, field_number)?.is_some();
    patch_fixed32_field(data, field_number, present, value)
}

const fn encode_i32(value: i32) -> u64 {
    value as i64 as u64
}

fn strict_optional_bool(data: &[u8], field_number: u32) -> Result<Option<bool>> {
    strict_optional_varint(data, field_number)?
        .map(|value| match value {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(Error::InvalidFormat(format!(
                "chart series error-bar field {field_number} is not a canonical boolean"
            ))),
        })
        .transpose()
}

fn strict_optional_i32(data: &[u8], field_number: u32) -> Result<Option<i32>> {
    strict_optional_varint(data, field_number)?
        .map(|value| {
            let decoded = value as i32;
            if encode_i32(decoded) != value && value > i32::MAX as u64 {
                return Err(Error::InvalidFormat(format!(
                    "chart series error-bar field {field_number} is not a canonical int32"
                )));
            }
            Ok(decoded)
        })
        .transpose()
}

fn strict_optional_f32(data: &[u8], field_number: u32) -> Result<Option<f32>> {
    strict_optional_fixed32(data, field_number)?
        .map(|bits| {
            let value = f32::from_bits(bits);
            if !value.is_finite() {
                return Err(Error::InvalidFormat(format!(
                    "chart series error-bar field {field_number} is not finite"
                )));
            }
            Ok(value)
        })
        .transpose()
}

fn strict_optional_fixed32(data: &[u8], field_number: u32) -> Result<Option<u32>> {
    let field = strict_optional_field(data, field_number)?;
    let Some(field) = field else {
        return Ok(None);
    };
    if field.wire_type() != 5 {
        return Err(Error::InvalidFormat(format!(
            "chart series error-bar field {field_number} is not fixed32"
        )));
    }
    let bytes: [u8; 4] = data[field.payload_start()..field.end()]
        .try_into()
        .map_err(|_| Error::InvalidFormat("truncated chart error-bar fixed32".to_owned()))?;
    Ok(Some(u32::from_le_bytes(bytes)))
}

fn strict_optional_bytes(data: &[u8], field_number: u32) -> Result<Option<&[u8]>> {
    let field = strict_optional_field(data, field_number)?;
    let Some(field) = field else {
        return Ok(None);
    };
    if field.wire_type() != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart series error-bar field {field_number} is not length-delimited"
        )));
    }
    Ok(Some(&data[field.payload_start()..field.end()]))
}

fn strict_optional_varint(data: &[u8], field_number: u32) -> Result<Option<u64>> {
    let field = strict_optional_field(data, field_number)?;
    let Some(field) = field else {
        return Ok(None);
    };
    if field.wire_type() != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart series error-bar field {field_number} is not a varint"
        )));
    }
    let (value, consumed) =
        litchi_iwa_common::varint::decode_varint_from_bytes(&data[field.key_end()..field.end()])
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "chart series error-bar field {field_number} is invalid: {error}"
                ))
            })?;
    if consumed != field.end() - field.key_end()
        || litchi_iwa_common::varint::encoded_len(value) != consumed
    {
        return Err(Error::InvalidFormat(format!(
            "chart series error-bar field {field_number} is not canonically encoded"
        )));
    }
    Ok(Some(value))
}

fn strict_optional_field(data: &[u8], field_number: u32) -> Result<Option<crate::wire::WireField>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields
        .into_iter()
        .filter(|field| field.number() == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart series error-bar field {field_number} occurs more than once"
        )));
    }
    Ok(Some(field))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::series_non_style::canonical_empty_chart_series_non_style_data;
    use crate::wire::append_varint_field;

    #[test]
    fn all_native_error_types_round_trip() {
        let direction = Direction::PositiveAndNegative;
        let values = CustomValues::new([1.0, 2.0], [0.5, 1.5]).unwrap();
        let variants = [
            Series::FixedValue {
                direction,
                value: FixedValue::new(12.5).unwrap(),
            },
            Series::Percentage {
                direction: Direction::PositiveOnly,
                percentage: Percentage::new(17).unwrap(),
            },
            Series::StandardDeviation {
                direction: Direction::NegativeOnly,
                deviations: StandardDeviationCount::new(3).unwrap(),
            },
            Series::StandardError { direction },
            Series::CustomValues { direction, values },
        ];
        for expected in variants {
            let original = canonical_empty_chart_series_non_style_data().unwrap();
            let patched = patch_error_bars(&original, &expected).unwrap();
            assert_eq!(read_error_bars(&patched).unwrap(), expected);
            let hidden = patch_error_bars(&patched, &Series::None).unwrap();
            assert_eq!(read_error_bars(&hidden).unwrap(), Series::None);
            assert_eq!(
                read_error_bars(&patch_error_bars(&hidden, &expected).unwrap()).unwrap(),
                expected
            );
        }
    }

    #[test]
    fn native_defaults_and_wire_validation_are_strict() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        assert_eq!(read_error_bars(&original).unwrap(), Series::None);
        let mut extension = Vec::new();
        append_varint_field(&mut extension, SHOW_ERROR_BARS_FIELD, 2).unwrap();
        let malformed =
            patch_chart_series_non_style_extension(&original, false, Some(&extension)).unwrap();
        assert!(read_error_bars(&malformed).is_err());
        assert!(FixedValue::new(0.0).is_err());
        assert!(Percentage::new(0).is_err());
        assert!(StandardDeviationCount::new(0).is_err());
        assert!(CustomValues::new([f64::INFINITY], []).is_err());
    }

    #[test]
    fn custom_value_patch_preserves_nested_unknown_fields() {
        const UNKNOWN_NESTED_FIELD: u32 = 99;
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let mut nested = Vec::new();
        append_varint_field(&mut nested, UNKNOWN_NESTED_FIELD, 73).unwrap();
        let extension =
            patch_length_delimited_field(&[], CUSTOM_POSITIVE_VALUES_FIELD, false, Some(&nested))
                .unwrap();
        let data =
            patch_chart_series_non_style_extension(&original, false, Some(&extension)).unwrap();
        let expected = Series::CustomValues {
            direction: Direction::PositiveOnly,
            values: CustomValues::new([2.5, 3.5], []).unwrap(),
        };
        let patched = patch_error_bars(&data, &expected).unwrap();
        let extension = generated_chart_series_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        let nested = strict_optional_bytes(extension, CUSTOM_POSITIVE_VALUES_FIELD)
            .unwrap()
            .unwrap();
        assert_eq!(
            strict_optional_varint(nested, UNKNOWN_NESTED_FIELD).unwrap(),
            Some(73)
        );
    }

    #[test]
    fn unknown_native_direction_and_value_type_round_trip_losslessly() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let expected = Series::unsupported(Direction::from_native(9_001), 9_002).unwrap();
        let patched = patch_error_bars(&original, &expected).unwrap();
        assert_eq!(read_error_bars(&patched).unwrap(), expected);
        let extension = generated_chart_series_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        assert_eq!(
            strict_optional_i32(extension, DIRECTION_FIELD).unwrap(),
            Some(9_001)
        );
        assert_eq!(
            strict_optional_i32(extension, VALUE_TYPE_FIELD).unwrap(),
            Some(9_002)
        );
        assert!(Series::unsupported(Direction::PositiveOnly, NATIVE_PERCENTAGE).is_err());
    }

    #[test]
    fn unknown_native_type_does_not_discard_unmodeled_value_fields() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let unknown = Series::unsupported(Direction::PositiveOnly, 9_002).unwrap();
        let mut extension = Vec::new();
        append_varint_field(&mut extension, SHOW_ERROR_BARS_FIELD, 1).unwrap();
        append_varint_field(&mut extension, DIRECTION_FIELD, 2).unwrap();
        append_varint_field(&mut extension, VALUE_TYPE_FIELD, 9_002).unwrap();
        let extension = patch_fixed32_field(
            &extension,
            FIXED_VALUE_FIELD,
            false,
            Some(12.5_f32.to_bits()),
        )
        .unwrap();
        let data =
            patch_chart_series_non_style_extension(&original, false, Some(&extension)).unwrap();
        let patched = patch_error_bars(&data, &unknown).unwrap();
        let extension = generated_chart_series_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        assert_eq!(
            strict_optional_fixed32(extension, FIXED_VALUE_FIELD).unwrap(),
            Some(12.5_f32.to_bits())
        );
        assert_eq!(read_error_bars(&patched).unwrap(), unknown);
    }
}
