//! Lossless per-series trendline-type CRUD for native charts.
//!
//! Pages, Numbers, and Keynote store an enabled trendline as two coordinated
//! fields in the generated series non-style extension: a visibility boolean
//! and a type discriminant. A native `None` selection omits both fields.

use prost::Message;

use crate::charts::series_non_style::{
    chart_series_non_style_values, generated_chart_series_non_style_extension,
    patch_chart_series_non_style_extension, set_chart_series_non_style_values,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

const SHOW_TRENDLINE_FIELD: u32 = 37;
const TRENDLINE_TYPE_FIELD: u32 = 62;

const NATIVE_NONE: i32 = 0;
const NATIVE_LINEAR: i32 = 1;
const NATIVE_LOGARITHMIC: i32 = 2;
const NATIVE_POLYNOMIAL: i32 = 3;
const NATIVE_POWER: i32 = 4;
const NATIVE_EXPONENTIAL: i32 = 5;
const NATIVE_MOVING_AVERAGE: i32 = 6;

/// The trendline fitted to one native chart series.
///
/// `Unsupported` preserves a future native type during read-modify-write
/// cycles. Constructing it with a currently known discriminant is rejected by
/// setters so one native value always has one canonical typed representation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum ChartSeriesTrendline {
    /// Do not draw a trendline.
    #[default]
    None,
    /// Fit a straight line.
    Linear,
    /// Fit a logarithmic curve.
    Logarithmic,
    /// Fit a polynomial curve. Its order is stored by a separate inspector control.
    Polynomial,
    /// Fit a power curve.
    Power,
    /// Fit an exponential curve.
    Exponential,
    /// Fit a moving average. Its period is stored by a separate inspector control.
    MovingAverage,
    /// Preserve an unrecognized native iWork trendline type.
    Unsupported(i32),
}

impl ChartSeriesTrendline {
    /// Decode the integer stored by the iWork protobuf schema.
    pub const fn from_raw(value: i32) -> Self {
        match value {
            NATIVE_NONE => Self::None,
            NATIVE_LINEAR => Self::Linear,
            NATIVE_LOGARITHMIC => Self::Logarithmic,
            NATIVE_POLYNOMIAL => Self::Polynomial,
            NATIVE_POWER => Self::Power,
            NATIVE_EXPONENTIAL => Self::Exponential,
            NATIVE_MOVING_AVERAGE => Self::MovingAverage,
            value => Self::Unsupported(value),
        }
    }

    /// Return the integer used by the iWork protobuf schema.
    pub const fn into_raw(self) -> i32 {
        match self {
            Self::None => NATIVE_NONE,
            Self::Linear => NATIVE_LINEAR,
            Self::Logarithmic => NATIVE_LOGARITHMIC,
            Self::Polynomial => NATIVE_POLYNOMIAL,
            Self::Power => NATIVE_POWER,
            Self::Exponential => NATIVE_EXPONENTIAL,
            Self::MovingAverage => NATIVE_MOVING_AVERAGE,
            Self::Unsupported(value) => value,
        }
    }

    /// Whether the series has a trendline.
    pub const fn is_visible(self) -> bool {
        !matches!(self, Self::None)
    }

    fn validate(self) -> Result<()> {
        if let Self::Unsupported(value) = self
            && matches!(
                value,
                NATIVE_NONE
                    | NATIVE_LINEAR
                    | NATIVE_LOGARITHMIC
                    | NATIVE_POLYNOMIAL
                    | NATIVE_POWER
                    | NATIVE_EXPONENTIAL
                    | NATIVE_MOVING_AVERAGE
            )
        {
            return Err(Error::InvalidFormat(format!(
                "known chart series trendline type {value} must use its named enum variant"
            )));
        }
        Ok(())
    }
}

/// Read trendline types in native series order.
pub(crate) fn chart_series_trendlines(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<Vec<ChartSeriesTrendline>> {
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        ChartSeriesTrendline::None,
        read_trendline,
    )
}

/// Set trendline types in native series order.
pub(crate) fn set_chart_series_trendlines(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    expected: &[ChartSeriesTrendline],
) -> Result<()> {
    for &trendline in expected {
        trendline.validate()?;
    }
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "series trendlines",
        expected,
        ChartSeriesTrendline::None,
        read_trendline,
        |data, trendline| patch_trendline(data, *trendline),
    )
}

fn read_trendline(data: &[u8]) -> Result<ChartSeriesTrendline> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(ChartSeriesTrendline::None);
    };
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let visible = strict_optional_bool(extension, SHOW_TRENDLINE_FIELD)?.unwrap_or(false);
    let native_type = strict_optional_i32(extension, TRENDLINE_TYPE_FIELD)?.unwrap_or(NATIVE_NONE);
    match (visible, native_type) {
        (false, NATIVE_NONE) => Ok(ChartSeriesTrendline::None),
        (true, NATIVE_NONE) => Err(Error::InvalidFormat(
            "visible chart series trendline has no type".to_owned(),
        )),
        (false, native_type) => Err(Error::InvalidFormat(format!(
            "hidden chart series trendline retains type {native_type}"
        ))),
        (true, native_type) => Ok(ChartSeriesTrendline::from_raw(native_type)),
    }
}

fn patch_trendline(data: &[u8], expected: ChartSeriesTrendline) -> Result<Vec<u8>> {
    expected.validate()?;
    let existing_extension = generated_chart_series_non_style_extension(data)?;
    if existing_extension.is_none() && expected == ChartSeriesTrendline::None {
        return Ok(data.to_vec());
    }
    let mut extension = existing_extension.unwrap_or_default().to_vec();
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension.as_slice())?;
    let show_present = strict_optional_bool(&extension, SHOW_TRENDLINE_FIELD)?.is_some();
    let type_present = strict_optional_i32(&extension, TRENDLINE_TYPE_FIELD)?.is_some();
    let visible = expected.is_visible();
    extension = patch_varint_field(
        &extension,
        SHOW_TRENDLINE_FIELD,
        show_present,
        visible.then_some(1),
    )?;
    extension = patch_varint_field(
        &extension,
        TRENDLINE_TYPE_FIELD,
        type_present,
        visible.then_some(encode_i32(expected.into_raw())),
    )?;
    let patched = patch_chart_series_non_style_extension(
        data,
        existing_extension.is_some(),
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    if read_trendline(&patched)? != expected {
        return Err(Error::InvalidFormat(
            "chart series trendline wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
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
                "chart series trendline field {field_number} is not a canonical boolean"
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
                    "chart series trendline field {field_number} is not a canonical int32"
                )));
            }
            Ok(decoded)
        })
        .transpose()
}

fn strict_optional_varint(data: &[u8], field_number: u32) -> Result<Option<u64>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart series trendline field {field_number} occurs more than once"
        )));
    }
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart series trendline field {field_number} is not a varint"
        )));
    }
    let (value, consumed) = crate::varint::decode_varint_from_bytes(
        &data[field.key_end..field.end],
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "chart series trendline field {field_number} is invalid: {error}"
        ))
    })?;
    if consumed != field.end - field.key_end {
        return Err(Error::InvalidFormat(format!(
            "chart series trendline field {field_number} has trailing bytes"
        )));
    }
    if data[field.key_end..field.end] != crate::varint::encode_varint(value) {
        return Err(Error::InvalidFormat(format!(
            "chart series trendline field {field_number} is not canonically encoded"
        )));
    }
    Ok(Some(value))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::series_non_style::canonical_empty_chart_series_non_style_data;
    use crate::wire::{append_varint_field, parse_wire_fields};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_EXTENSION_FIELD: u32 = 4_097;

    #[test]
    fn every_native_inspector_type_round_trips() {
        let variants = [
            ChartSeriesTrendline::Linear,
            ChartSeriesTrendline::Logarithmic,
            ChartSeriesTrendline::Polynomial,
            ChartSeriesTrendline::Power,
            ChartSeriesTrendline::Exponential,
            ChartSeriesTrendline::MovingAverage,
        ];
        for trendline in variants {
            let original = canonical_empty_chart_series_non_style_data().unwrap();
            let patched = patch_trendline(&original, trendline).unwrap();
            assert_eq!(read_trendline(&patched).unwrap(), trendline);
            assert_eq!(
                patch_trendline(&patched, ChartSeriesTrendline::None).unwrap(),
                original
            );
        }
    }

    #[test]
    fn patch_preserves_neighboring_and_unknown_fields() {
        let mut extension = Vec::new();
        append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, 91).unwrap();
        let mut original = canonical_empty_chart_series_non_style_data().unwrap();
        original =
            patch_chart_series_non_style_extension(&original, false, Some(extension.as_slice()))
                .unwrap();
        append_varint_field(&mut original, UNKNOWN_OUTER_FIELD, 73).unwrap();

        let patched = patch_trendline(&original, ChartSeriesTrendline::Power).unwrap();
        assert_eq!(
            parse_wire_fields(&patched)
                .unwrap()
                .iter()
                .find(|field| field.number == UNKNOWN_OUTER_FIELD)
                .map(|field| &patched[field.start..field.end]),
            parse_wire_fields(&original)
                .unwrap()
                .iter()
                .find(|field| field.number == UNKNOWN_OUTER_FIELD)
                .map(|field| &original[field.start..field.end])
        );
        let extension = generated_chart_series_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        assert_eq!(
            strict_optional_varint(extension, UNKNOWN_EXTENSION_FIELD).unwrap(),
            Some(91)
        );
    }

    #[test]
    fn inconsistent_or_noncanonical_native_fields_are_rejected() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let mut extension = Vec::new();
        append_varint_field(&mut extension, SHOW_TRENDLINE_FIELD, 1).unwrap();
        let missing_type =
            patch_chart_series_non_style_extension(&original, false, Some(&extension)).unwrap();
        assert!(read_trendline(&missing_type).is_err());

        extension.clear();
        append_varint_field(&mut extension, TRENDLINE_TYPE_FIELD, NATIVE_LINEAR as u64).unwrap();
        let hidden_type =
            patch_chart_series_non_style_extension(&original, false, Some(&extension)).unwrap();
        assert!(read_trendline(&hidden_type).is_err());

        extension.clear();
        append_varint_field(&mut extension, SHOW_TRENDLINE_FIELD, 2).unwrap();
        append_varint_field(&mut extension, TRENDLINE_TYPE_FIELD, NATIVE_LINEAR as u64).unwrap();
        let malformed_bool =
            patch_chart_series_non_style_extension(&original, false, Some(&extension)).unwrap();
        assert!(read_trendline(&malformed_bool).is_err());

        extension.clear();
        extension.extend(crate::varint::encode_varint(
            u64::from(SHOW_TRENDLINE_FIELD) << 3,
        ));
        extension.extend([0x81, 0x00]);
        append_varint_field(&mut extension, TRENDLINE_TYPE_FIELD, NATIVE_LINEAR as u64).unwrap();
        let overlong_bool =
            patch_chart_series_non_style_extension(&original, false, Some(&extension)).unwrap();
        assert!(read_trendline(&overlong_bool).is_err());
    }

    #[test]
    fn unsupported_values_are_lossless_but_known_aliases_are_rejected() {
        const FUTURE_TYPE: i32 = 9_001;
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let future = ChartSeriesTrendline::Unsupported(FUTURE_TYPE);
        let patched = patch_trendline(&original, future).unwrap();
        assert_eq!(read_trendline(&patched).unwrap(), future);
        assert!(
            patch_trendline(&original, ChartSeriesTrendline::Unsupported(NATIVE_LINEAR)).is_err()
        );
    }
}
