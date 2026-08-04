//! Lossless per-series error-bar CRUD for native charts.
//!
//! The native inspector coordinates visibility, direction, value derivation,
//! and type-specific values inside each series non-style extension.

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
use prost::Message;

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
const NATIVE_DIRECTION_POSITIVE: i32 = 2;
const NATIVE_DIRECTION_NEGATIVE: i32 = 3;

const NATIVE_FIXED_VALUE: i32 = 1;
const NATIVE_PERCENTAGE: i32 = 2;
const NATIVE_STANDARD_DEVIATION: i32 = 3;
const NATIVE_STANDARD_ERROR: i32 = 4;
const NATIVE_CUSTOM_VALUES: i32 = 5;

const DEFAULT_FIXED_VALUE: f32 = 10.0;
const DEFAULT_PERCENTAGE: u8 = 5;
const DEFAULT_STANDARD_DEVIATIONS: u32 = 1;
const MAX_EXACT_F32_INTEGER: u32 = 1 << f32::MANTISSA_DIGITS;
const MAX_CUSTOM_VALUES_PER_SIDE: usize = 1_000_000;

/// Which side of each data point receives an error bar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartErrorBarDirection {
    /// Draw error bars above and below each value.
    PositiveAndNegative,
    /// Draw only the positive error bar.
    PositiveOnly,
    /// Draw only the negative error bar.
    NegativeOnly,
    /// Preserve an unrecognized native direction.
    Unsupported(i32),
}

impl ChartErrorBarDirection {
    const fn from_raw(value: i32) -> Self {
        match value {
            NATIVE_DIRECTION_BOTH => Self::PositiveAndNegative,
            NATIVE_DIRECTION_POSITIVE => Self::PositiveOnly,
            NATIVE_DIRECTION_NEGATIVE => Self::NegativeOnly,
            value => Self::Unsupported(value),
        }
    }

    const fn into_raw(self) -> i32 {
        match self {
            Self::PositiveAndNegative => NATIVE_DIRECTION_BOTH,
            Self::PositiveOnly => NATIVE_DIRECTION_POSITIVE,
            Self::NegativeOnly => NATIVE_DIRECTION_NEGATIVE,
            Self::Unsupported(value) => value,
        }
    }

    fn validate(self) -> Result<()> {
        if let Self::Unsupported(value) = self
            && matches!(
                value,
                NATIVE_DIRECTION_BOTH | NATIVE_DIRECTION_POSITIVE | NATIVE_DIRECTION_NEGATIVE
            )
        {
            return Err(Error::InvalidFormat(format!(
                "known chart error-bar direction {value} must use its named variant"
            )));
        }
        Ok(())
    }
}

/// Positive finite fixed error magnitude.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartErrorBarFixedValue(f32);

impl ChartErrorBarFixedValue {
    /// Native inspector default.
    pub const DEFAULT: Self = Self(DEFAULT_FIXED_VALUE);

    /// Validate and construct a fixed error magnitude.
    pub fn new(value: f32) -> Result<Self> {
        if !value.is_finite() || value <= 0.0 {
            return Err(Error::InvalidFormat(format!(
                "chart fixed error-bar value must be positive and finite, got {value}"
            )));
        }
        Ok(Self(value))
    }

    /// Return the validated magnitude.
    pub const fn get(self) -> f32 {
        self.0
    }
}

impl Default for ChartErrorBarFixedValue {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Percentage accepted by the iWork error-bar inspector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartErrorBarPercentage(u8);

impl ChartErrorBarPercentage {
    /// Smallest accepted percentage.
    pub const MIN: u8 = 1;
    /// Largest accepted percentage.
    pub const MAX: u8 = 100;
    /// Native inspector default.
    pub const DEFAULT: Self = Self(DEFAULT_PERCENTAGE);

    /// Validate and construct an error percentage.
    pub fn new(value: u8) -> Result<Self> {
        if !(Self::MIN..=Self::MAX).contains(&value) {
            return Err(Error::InvalidFormat(format!(
                "chart error-bar percentage must be {}..={}, got {value}",
                Self::MIN,
                Self::MAX
            )));
        }
        Ok(Self(value))
    }

    /// Return the validated percentage.
    pub const fn get(self) -> u8 {
        self.0
    }
}

impl Default for ChartErrorBarPercentage {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Positive integral count used by standard-deviation error bars.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartErrorBarStandardDeviationCount(u32);

impl ChartErrorBarStandardDeviationCount {
    /// Smallest accepted count.
    pub const MIN: u32 = 1;
    /// Largest consecutive integer represented exactly by native `float`.
    pub const MAX: u32 = MAX_EXACT_F32_INTEGER;
    /// Native inspector default.
    pub const DEFAULT: Self = Self(DEFAULT_STANDARD_DEVIATIONS);

    /// Validate and construct a standard-deviation count.
    pub fn new(value: u32) -> Result<Self> {
        if !(Self::MIN..=Self::MAX).contains(&value) {
            return Err(Error::InvalidFormat(format!(
                "chart error-bar standard-deviation count must be {}..={}, got {value}",
                Self::MIN,
                Self::MAX
            )));
        }
        Ok(Self(value))
    }

    /// Return the validated count.
    pub const fn get(self) -> u32 {
        self.0
    }
}

impl Default for ChartErrorBarStandardDeviationCount {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// One finite, nonnegative custom error magnitude.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartErrorBarCustomValue(f64);

impl ChartErrorBarCustomValue {
    /// Validate and construct a custom error magnitude.
    pub fn new(value: f64) -> Result<Self> {
        if !value.is_finite() || value < 0.0 {
            return Err(Error::InvalidFormat(format!(
                "custom chart error-bar value must be nonnegative and finite, got {value}"
            )));
        }
        Ok(Self(value))
    }

    /// Return the validated magnitude.
    pub const fn get(self) -> f64 {
        self.0
    }
}

/// Per-point custom positive and negative error magnitudes.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartErrorBarCustomValues {
    positive: Box<[ChartErrorBarCustomValue]>,
    negative: Box<[ChartErrorBarCustomValue]>,
}

impl ChartErrorBarCustomValues {
    /// Validate and collect custom positive and negative magnitudes.
    pub fn new(
        positive: impl IntoIterator<Item = f64>,
        negative: impl IntoIterator<Item = f64>,
    ) -> Result<Self> {
        Ok(Self {
            positive: collect_custom_values(positive, "positive")?,
            negative: collect_custom_values(negative, "negative")?,
        })
    }

    /// Positive custom magnitudes in native point order.
    pub fn positive(&self) -> &[ChartErrorBarCustomValue] {
        &self.positive
    }

    /// Negative custom magnitudes in native point order.
    pub fn negative(&self) -> &[ChartErrorBarCustomValue] {
        &self.negative
    }
}

/// Complete error-bar configuration for one native chart series.
#[derive(Debug, Clone, PartialEq, Default)]
#[non_exhaustive]
pub enum ChartSeriesErrorBars {
    /// Do not draw error bars.
    #[default]
    None,
    /// Use one fixed magnitude for every point.
    FixedValue {
        /// Sides on which bars are drawn.
        direction: ChartErrorBarDirection,
        /// Fixed error magnitude.
        value: ChartErrorBarFixedValue,
    },
    /// Derive each error from a percentage of its point value.
    Percentage {
        /// Sides on which bars are drawn.
        direction: ChartErrorBarDirection,
        /// Percentage of each point value.
        percentage: ChartErrorBarPercentage,
    },
    /// Derive errors from an integral number of standard deviations.
    StandardDeviation {
        /// Sides on which bars are drawn.
        direction: ChartErrorBarDirection,
        /// Number of standard deviations.
        deviations: ChartErrorBarStandardDeviationCount,
    },
    /// Derive errors from the standard error.
    StandardError {
        /// Sides on which bars are drawn.
        direction: ChartErrorBarDirection,
    },
    /// Use custom positive and negative magnitudes in native point order.
    CustomValues {
        /// Sides on which bars are drawn.
        direction: ChartErrorBarDirection,
        /// Per-point magnitudes.
        values: ChartErrorBarCustomValues,
    },
    /// Preserve an unrecognized native value-derivation type.
    Unsupported {
        /// Sides on which bars are drawn.
        direction: ChartErrorBarDirection,
        /// Unrecognized native type discriminant.
        native_type: i32,
    },
}

impl ChartSeriesErrorBars {
    /// Native inspector defaults for fixed-value errors.
    pub const fn fixed_value(direction: ChartErrorBarDirection) -> Self {
        Self::FixedValue {
            direction,
            value: ChartErrorBarFixedValue::DEFAULT,
        }
    }

    /// Native inspector defaults for percentage errors.
    pub const fn percentage(direction: ChartErrorBarDirection) -> Self {
        Self::Percentage {
            direction,
            percentage: ChartErrorBarPercentage::DEFAULT,
        }
    }

    /// Native inspector defaults for standard-deviation errors.
    pub const fn standard_deviation(direction: ChartErrorBarDirection) -> Self {
        Self::StandardDeviation {
            direction,
            deviations: ChartErrorBarStandardDeviationCount::DEFAULT,
        }
    }

    fn direction(&self) -> Option<ChartErrorBarDirection> {
        match self {
            Self::None => None,
            Self::FixedValue { direction, .. }
            | Self::Percentage { direction, .. }
            | Self::StandardDeviation { direction, .. }
            | Self::StandardError { direction }
            | Self::CustomValues { direction, .. }
            | Self::Unsupported { direction, .. } => Some(*direction),
        }
    }

    fn native_type(&self) -> Option<i32> {
        match self {
            Self::None => None,
            Self::FixedValue { .. } => Some(NATIVE_FIXED_VALUE),
            Self::Percentage { .. } => Some(NATIVE_PERCENTAGE),
            Self::StandardDeviation { .. } => Some(NATIVE_STANDARD_DEVIATION),
            Self::StandardError { .. } => Some(NATIVE_STANDARD_ERROR),
            Self::CustomValues { .. } => Some(NATIVE_CUSTOM_VALUES),
            Self::Unsupported { native_type, .. } => Some(*native_type),
        }
    }

    fn validate(&self) -> Result<()> {
        let Some(direction) = self.direction() else {
            return Ok(());
        };
        direction.validate()?;
        if let Self::Unsupported { native_type, .. } = self
            && matches!(
                *native_type,
                NATIVE_FIXED_VALUE
                    | NATIVE_PERCENTAGE
                    | NATIVE_STANDARD_DEVIATION
                    | NATIVE_STANDARD_ERROR
                    | NATIVE_CUSTOM_VALUES
            )
        {
            return Err(Error::InvalidFormat(format!(
                "known chart error-bar value type {native_type} must use its named variant"
            )));
        }
        Ok(())
    }
}

pub(crate) fn chart_series_error_bars(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<Vec<ChartSeriesErrorBars>> {
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        ChartSeriesErrorBars::None,
        read_error_bars,
    )
}

pub(crate) fn set_chart_series_error_bars(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    expected: &[ChartSeriesErrorBars],
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
        ChartSeriesErrorBars::None,
        read_error_bars,
        patch_error_bars,
    )
}

fn read_error_bars(data: &[u8]) -> Result<ChartSeriesErrorBars> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(ChartSeriesErrorBars::None);
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
        return Ok(ChartSeriesErrorBars::None);
    }

    let direction = ChartErrorBarDirection::from_raw(direction);
    direction.validate()?;
    let value = match native_type {
        NATIVE_FIXED_VALUE => ChartSeriesErrorBars::FixedValue {
            direction,
            value: ChartErrorBarFixedValue::new(fixed.unwrap_or(DEFAULT_FIXED_VALUE))?,
        },
        NATIVE_PERCENTAGE => ChartSeriesErrorBars::Percentage {
            direction,
            percentage: ChartErrorBarPercentage::new(f32_to_u8(
                percentage.unwrap_or(f32::from(DEFAULT_PERCENTAGE)),
                "percentage",
            )?)?,
        },
        NATIVE_STANDARD_DEVIATION => ChartSeriesErrorBars::StandardDeviation {
            direction,
            deviations: ChartErrorBarStandardDeviationCount::new(f32_to_u32(
                deviations.unwrap_or(DEFAULT_STANDARD_DEVIATIONS as f32),
                "standard-deviation count",
            )?)?,
        },
        NATIVE_STANDARD_ERROR => ChartSeriesErrorBars::StandardError { direction },
        NATIVE_CUSTOM_VALUES => ChartSeriesErrorBars::CustomValues {
            direction,
            values: ChartErrorBarCustomValues {
                positive: positive.into_boxed_slice(),
                negative: negative.into_boxed_slice(),
            },
        },
        native_type => ChartSeriesErrorBars::Unsupported {
            direction,
            native_type,
        },
    };
    value.validate()?;
    Ok(value)
}

fn patch_error_bars(data: &[u8], expected: &ChartSeriesErrorBars) -> Result<Vec<u8>> {
    expected.validate()?;
    let existing_extension = generated_chart_series_non_style_extension(data)?;
    if existing_extension.is_none() && expected == &ChartSeriesErrorBars::None {
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
            Some(encode_i32(direction.into_raw())),
        )?;
        extension = patch_optional_varint(
            &extension,
            VALUE_TYPE_FIELD,
            expected.native_type().map(encode_i32),
        )?;
        match expected {
            ChartSeriesErrorBars::FixedValue { value, .. } => {
                extension = patch_optional_fixed32(
                    &extension,
                    FIXED_VALUE_FIELD,
                    Some(value.get().to_bits()),
                )?;
            },
            ChartSeriesErrorBars::Percentage { percentage, .. } => {
                extension = patch_optional_fixed32(
                    &extension,
                    PERCENTAGE_FIELD,
                    Some(f32::from(percentage.get()).to_bits()),
                )?;
            },
            ChartSeriesErrorBars::StandardDeviation { deviations, .. } => {
                extension = patch_optional_fixed32(
                    &extension,
                    STANDARD_DEVIATION_FIELD,
                    Some((deviations.get() as f32).to_bits()),
                )?;
            },
            ChartSeriesErrorBars::CustomValues { values, .. } => {
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
            ChartSeriesErrorBars::None
            | ChartSeriesErrorBars::StandardError { .. }
            | ChartSeriesErrorBars::Unsupported { .. } => {},
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
    expected: &[ChartErrorBarCustomValue],
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
            .map(|value| value.get().to_bits())
            .collect::<Vec<_>>(),
    )?;
    patch_length_delimited_field(
        extension,
        field_number,
        existing.is_some(),
        (!nested.is_empty()).then_some(nested.as_slice()),
    )
}

fn read_custom_values(
    extension: &[u8],
    field_number: u32,
) -> Result<Vec<ChartErrorBarCustomValue>> {
    let Some(nested) = strict_optional_bytes(extension, field_number)? else {
        return Ok(Vec::new());
    };
    tsch::ChartsNsArrayOfNsNumberDoubleArchive::decode(nested)?;
    repeated_fixed64_values(nested, CUSTOM_ARRAY_VALUE_FIELD)?
        .into_iter()
        .map(|bits| ChartErrorBarCustomValue::new(f64::from_bits(bits)))
        .collect()
}

fn collect_custom_values(
    values: impl IntoIterator<Item = f64>,
    side: &str,
) -> Result<Box<[ChartErrorBarCustomValue]>> {
    let mut collected = Vec::new();
    for value in values {
        if collected.len() == MAX_CUSTOM_VALUES_PER_SIDE {
            return Err(Error::InvalidFormat(format!(
                "custom chart error-bar {side} values exceed {MAX_CUSTOM_VALUES_PER_SIDE}"
            )));
        }
        collected.push(ChartErrorBarCustomValue::new(value)?);
    }
    Ok(collected.into_boxed_slice())
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
        let direction = ChartErrorBarDirection::PositiveAndNegative;
        let values = ChartErrorBarCustomValues::new([1.0, 2.0], [0.5, 1.5]).unwrap();
        let variants = [
            ChartSeriesErrorBars::FixedValue {
                direction,
                value: ChartErrorBarFixedValue::new(12.5).unwrap(),
            },
            ChartSeriesErrorBars::Percentage {
                direction: ChartErrorBarDirection::PositiveOnly,
                percentage: ChartErrorBarPercentage::new(17).unwrap(),
            },
            ChartSeriesErrorBars::StandardDeviation {
                direction: ChartErrorBarDirection::NegativeOnly,
                deviations: ChartErrorBarStandardDeviationCount::new(3).unwrap(),
            },
            ChartSeriesErrorBars::StandardError { direction },
            ChartSeriesErrorBars::CustomValues { direction, values },
        ];
        for expected in variants {
            let original = canonical_empty_chart_series_non_style_data().unwrap();
            let patched = patch_error_bars(&original, &expected).unwrap();
            assert_eq!(read_error_bars(&patched).unwrap(), expected);
            let hidden = patch_error_bars(&patched, &ChartSeriesErrorBars::None).unwrap();
            assert_eq!(
                read_error_bars(&hidden).unwrap(),
                ChartSeriesErrorBars::None
            );
            assert_eq!(
                read_error_bars(&patch_error_bars(&hidden, &expected).unwrap()).unwrap(),
                expected
            );
        }
    }

    #[test]
    fn native_defaults_and_public_bounds_are_strict() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        assert_eq!(
            read_error_bars(&original).unwrap(),
            ChartSeriesErrorBars::None
        );
        assert!(ChartErrorBarFixedValue::new(0.0).is_err());
        assert!(ChartErrorBarFixedValue::new(f32::NAN).is_err());
        assert!(ChartErrorBarPercentage::new(0).is_err());
        assert!(ChartErrorBarPercentage::new(101).is_err());
        assert!(ChartErrorBarStandardDeviationCount::new(0).is_err());
        assert!(
            ChartErrorBarStandardDeviationCount::new(ChartErrorBarStandardDeviationCount::MAX + 1)
                .is_err()
        );
        assert!(ChartErrorBarCustomValues::new([f64::INFINITY], []).is_err());
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
        let expected = ChartSeriesErrorBars::CustomValues {
            direction: ChartErrorBarDirection::PositiveOnly,
            values: ChartErrorBarCustomValues::new([2.5, 3.5], []).unwrap(),
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
    fn malformed_and_known_alias_values_are_rejected() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let mut extension = Vec::new();
        append_varint_field(&mut extension, SHOW_ERROR_BARS_FIELD, 2).unwrap();
        let malformed =
            patch_chart_series_non_style_extension(&original, false, Some(&extension)).unwrap();
        assert!(read_error_bars(&malformed).is_err());

        let known_alias = ChartSeriesErrorBars::Unsupported {
            direction: ChartErrorBarDirection::PositiveOnly,
            native_type: NATIVE_PERCENTAGE,
        };
        assert!(patch_error_bars(&original, &known_alias).is_err());
        let aliased_direction = ChartSeriesErrorBars::StandardError {
            direction: ChartErrorBarDirection::Unsupported(NATIVE_DIRECTION_BOTH),
        };
        assert!(patch_error_bars(&original, &aliased_direction).is_err());
    }
}
