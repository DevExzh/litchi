//! Lossless per-series trendline CRUD for native charts.
//!
//! Pages, Numbers, and Keynote keep trendline type, fitting parameters, legend
//! label, equation visibility, and R² visibility in the generated series
//! non-style extension. This module treats those coordinated fields as one
//! validated value.

use prost::Message;

use crate::charts::series_non_style::{
    NewChartSeriesNonStyleBase, chart_series_non_style_values,
    generated_chart_series_non_style_extension, patch_chart_series_non_style_extension,
    set_chart_series_non_style_values,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

const SHOW_TRENDLINE_FIELD: u32 = 37;
const TRENDLINE_LABEL_FIELD: u32 = 54;
const POLYNOMIAL_ORDER_FIELD: u32 = 55;
const MOVING_AVERAGE_PERIOD_FIELD: u32 = 56;
const SHOW_EQUATION_FIELD: u32 = 59;
const SHOW_LABEL_FIELD: u32 = 60;
const SHOW_R_SQUARED_FIELD: u32 = 61;
const TRENDLINE_TYPE_FIELD: u32 = 62;

const NATIVE_NONE: i32 = 0;
const NATIVE_LINEAR: i32 = 1;
const NATIVE_LOGARITHMIC: i32 = 2;
const NATIVE_POLYNOMIAL: i32 = 3;
const NATIVE_POWER: i32 = 4;
const NATIVE_EXPONENTIAL: i32 = 5;
const NATIVE_MOVING_AVERAGE: i32 = 6;

/// Polynomial order accepted by the iWork trendline inspector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartSeriesTrendlinePolynomialOrder(u8);

impl ChartSeriesTrendlinePolynomialOrder {
    /// Smallest polynomial order accepted by Pages, Numbers, and Keynote.
    pub const MIN: u8 = 2;
    /// Largest polynomial order accepted by Pages, Numbers, and Keynote.
    pub const MAX: u8 = 6;
    /// Native inspector default.
    pub const DEFAULT: Self = Self(Self::MIN);

    /// Validate and construct a polynomial order.
    pub fn new(value: u8) -> Result<Self> {
        if !(Self::MIN..=Self::MAX).contains(&value) {
            return Err(Error::InvalidFormat(format!(
                "chart series trendline polynomial order must be {}..={}, got {value}",
                Self::MIN,
                Self::MAX
            )));
        }
        Ok(Self(value))
    }

    /// Return the validated order.
    pub const fn get(self) -> u8 {
        self.0
    }
}

impl Default for ChartSeriesTrendlinePolynomialOrder {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Moving-average period accepted by the iWork trendline inspector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartSeriesTrendlineMovingAveragePeriod(u32);

impl ChartSeriesTrendlineMovingAveragePeriod {
    /// Smallest period accepted by Pages, Numbers, and Keynote.
    pub const MIN: u32 = 2;
    /// Largest value representable by the native signed integer field.
    pub const MAX: u32 = i32::MAX as u32;
    /// Native inspector default.
    pub const DEFAULT: Self = Self(Self::MIN);

    /// Validate and construct a moving-average period.
    pub fn new(value: u32) -> Result<Self> {
        if !(Self::MIN..=Self::MAX).contains(&value) {
            return Err(Error::InvalidFormat(format!(
                "chart series trendline moving-average period must be {}..={}, got {value}",
                Self::MIN,
                Self::MAX
            )));
        }
        Ok(Self(value))
    }

    /// Return the validated period.
    pub const fn get(self) -> u32 {
        self.0
    }
}

impl Default for ChartSeriesTrendlineMovingAveragePeriod {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Curve family and its type-specific fitting parameter.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum ChartSeriesTrendlineType {
    /// Do not draw a trendline.
    #[default]
    None,
    /// Fit a straight line.
    Linear,
    /// Fit a logarithmic curve.
    Logarithmic,
    /// Fit a polynomial curve with the given validated order.
    Polynomial(ChartSeriesTrendlinePolynomialOrder),
    /// Fit a power curve.
    Power,
    /// Fit an exponential curve.
    Exponential,
    /// Fit a moving average with the given validated period.
    MovingAverage(ChartSeriesTrendlineMovingAveragePeriod),
    /// Preserve an unrecognized native iWork trendline type.
    Unsupported(i32),
}

impl ChartSeriesTrendlineType {
    const fn native_type(self) -> i32 {
        match self {
            Self::None => NATIVE_NONE,
            Self::Linear => NATIVE_LINEAR,
            Self::Logarithmic => NATIVE_LOGARITHMIC,
            Self::Polynomial(_) => NATIVE_POLYNOMIAL,
            Self::Power => NATIVE_POWER,
            Self::Exponential => NATIVE_EXPONENTIAL,
            Self::MovingAverage(_) => NATIVE_MOVING_AVERAGE,
            Self::Unsupported(value) => value,
        }
    }

    const fn is_visible(self) -> bool {
        !matches!(self, Self::None)
    }

    const fn supports_equation(self) -> bool {
        !matches!(self, Self::None | Self::MovingAverage(_))
    }
}

/// Complete trendline configuration for one native chart series.
///
/// Use the named constructors for curve-specific defaults, then opt into
/// legend, equation, or R² display through the checked builder methods.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct ChartSeriesTrendline {
    trendline_type: ChartSeriesTrendlineType,
    custom_name: Option<String>,
    show_name_in_legend: bool,
    show_equation: bool,
    show_r_squared: bool,
}

impl ChartSeriesTrendline {
    /// No trendline.
    pub const fn none() -> Self {
        Self {
            trendline_type: ChartSeriesTrendlineType::None,
            custom_name: None,
            show_name_in_legend: false,
            show_equation: false,
            show_r_squared: false,
        }
    }

    /// A linear trendline.
    pub const fn linear() -> Self {
        Self::from_type(ChartSeriesTrendlineType::Linear)
    }

    /// A logarithmic trendline.
    pub const fn logarithmic() -> Self {
        Self::from_type(ChartSeriesTrendlineType::Logarithmic)
    }

    /// A polynomial trendline with a validated order.
    pub const fn polynomial(order: ChartSeriesTrendlinePolynomialOrder) -> Self {
        Self::from_type(ChartSeriesTrendlineType::Polynomial(order))
    }

    /// A power trendline.
    pub const fn power() -> Self {
        Self::from_type(ChartSeriesTrendlineType::Power)
    }

    /// An exponential trendline.
    pub const fn exponential() -> Self {
        Self::from_type(ChartSeriesTrendlineType::Exponential)
    }

    /// A moving-average trendline with a validated period.
    pub const fn moving_average(period: ChartSeriesTrendlineMovingAveragePeriod) -> Self {
        Self::from_type(ChartSeriesTrendlineType::MovingAverage(period))
    }

    /// Preserve a future native trendline type.
    pub fn unsupported(native_type: i32) -> Result<Self> {
        let value = Self::from_type(ChartSeriesTrendlineType::Unsupported(native_type));
        value.validate()?;
        Ok(value)
    }

    const fn from_type(trendline_type: ChartSeriesTrendlineType) -> Self {
        Self {
            trendline_type,
            custom_name: None,
            show_name_in_legend: false,
            show_equation: false,
            show_r_squared: false,
        }
    }

    /// Curve family and its type-specific fitting parameter.
    pub const fn trendline_type(&self) -> ChartSeriesTrendlineType {
        self.trendline_type
    }

    /// Custom legend name, if one replaces the automatic name.
    pub fn custom_name(&self) -> Option<&str> {
        self.custom_name.as_deref()
    }

    /// Whether the trendline name is shown in the chart legend.
    pub const fn shows_name_in_legend(&self) -> bool {
        self.show_name_in_legend
    }

    /// Whether the fitted equation is shown on the chart.
    pub const fn shows_equation(&self) -> bool {
        self.show_equation
    }

    /// Whether the fitted R² value is shown on the chart.
    pub const fn shows_r_squared(&self) -> bool {
        self.show_r_squared
    }

    /// Show or hide the automatic trendline name in the legend.
    pub fn with_legend_visibility(mut self, visible: bool) -> Result<Self> {
        self.show_name_in_legend = visible;
        self.validate()?;
        Ok(self)
    }

    /// Use and show a custom trendline name in the legend.
    pub fn with_legend_name(mut self, name: impl Into<String>) -> Result<Self> {
        self.custom_name = Some(name.into());
        self.show_name_in_legend = true;
        self.validate()?;
        Ok(self)
    }

    /// Remove a custom name while retaining the current legend visibility.
    pub fn without_custom_name(mut self) -> Self {
        self.custom_name = None;
        self
    }

    /// Show or hide the fitted equation.
    pub fn with_equation_visibility(mut self, visible: bool) -> Result<Self> {
        self.show_equation = visible;
        self.validate()?;
        Ok(self)
    }

    /// Show or hide the fitted R² value.
    pub fn with_r_squared_visibility(mut self, visible: bool) -> Result<Self> {
        self.show_r_squared = visible;
        self.validate()?;
        Ok(self)
    }

    fn validate(&self) -> Result<()> {
        if let ChartSeriesTrendlineType::Unsupported(value) = self.trendline_type
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
                "known chart series trendline type {value} must use its named representation"
            )));
        }
        if !self.trendline_type.is_visible()
            && (self.custom_name.is_some()
                || self.show_name_in_legend
                || self.show_equation
                || self.show_r_squared)
        {
            return Err(Error::InvalidFormat(
                "a hidden chart series trendline cannot expose display options".to_owned(),
            ));
        }
        if !self.trendline_type.supports_equation() && (self.show_equation || self.show_r_squared) {
            return Err(Error::InvalidFormat(
                "moving-average trendlines do not support equation or R² labels".to_owned(),
            ));
        }
        Ok(())
    }
}

/// Read complete trendline configurations in native series order.
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
        ChartSeriesTrendline::none(),
        read_trendline,
    )
}

/// Set complete trendline configurations in native series order.
pub(crate) fn set_chart_series_trendlines(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    expected: &[ChartSeriesTrendline],
) -> Result<()> {
    for trendline in expected {
        trendline.validate()?;
    }
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "series trendlines",
        NewChartSeriesNonStyleBase::Styled,
        expected,
        ChartSeriesTrendline::none(),
        read_trendline,
        patch_trendline,
    )
}

fn read_trendline(data: &[u8]) -> Result<ChartSeriesTrendline> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(ChartSeriesTrendline::none());
    };
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let visible = strict_optional_bool(extension, SHOW_TRENDLINE_FIELD)?.unwrap_or(false);
    let native_type = strict_optional_i32(extension, TRENDLINE_TYPE_FIELD)?.unwrap_or(NATIVE_NONE);
    if !visible && native_type == NATIVE_NONE {
        return Ok(ChartSeriesTrendline::none());
    }
    if visible && native_type == NATIVE_NONE {
        return Err(Error::InvalidFormat(
            "visible chart series trendline has no type".to_owned(),
        ));
    }
    if !visible {
        return Err(Error::InvalidFormat(format!(
            "hidden chart series trendline retains type {native_type}"
        )));
    }

    let polynomial_order = strict_optional_i32(extension, POLYNOMIAL_ORDER_FIELD)?;
    let moving_average_period = strict_optional_i32(extension, MOVING_AVERAGE_PERIOD_FIELD)?;
    let native_show_equation = strict_optional_bool(extension, SHOW_EQUATION_FIELD)?;
    let native_show_r_squared = strict_optional_bool(extension, SHOW_R_SQUARED_FIELD)?;
    let trendline_type = match native_type {
        NATIVE_LINEAR => ChartSeriesTrendlineType::Linear,
        NATIVE_LOGARITHMIC => ChartSeriesTrendlineType::Logarithmic,
        NATIVE_POLYNOMIAL => {
            let raw =
                polynomial_order.unwrap_or(i32::from(ChartSeriesTrendlinePolynomialOrder::MIN));
            let value = u8::try_from(raw).map_err(|_| {
                Error::InvalidFormat(format!(
                    "chart series trendline polynomial order is out of range: {raw}"
                ))
            })?;
            ChartSeriesTrendlineType::Polynomial(ChartSeriesTrendlinePolynomialOrder::new(value)?)
        },
        NATIVE_POWER => ChartSeriesTrendlineType::Power,
        NATIVE_EXPONENTIAL => ChartSeriesTrendlineType::Exponential,
        NATIVE_MOVING_AVERAGE => {
            let raw = moving_average_period
                .unwrap_or(ChartSeriesTrendlineMovingAveragePeriod::MIN as i32);
            let value = u32::try_from(raw).map_err(|_| {
                Error::InvalidFormat(format!(
                    "chart series trendline moving-average period is out of range: {raw}"
                ))
            })?;
            ChartSeriesTrendlineType::MovingAverage(ChartSeriesTrendlineMovingAveragePeriod::new(
                value,
            )?)
        },
        value => ChartSeriesTrendlineType::Unsupported(value),
    };
    let supports_equation = trendline_type.supports_equation();
    let value = ChartSeriesTrendline {
        trendline_type,
        custom_name: strict_optional_string(extension, TRENDLINE_LABEL_FIELD)?,
        show_name_in_legend: strict_optional_bool(extension, SHOW_LABEL_FIELD)?.unwrap_or(false),
        show_equation: supports_equation && native_show_equation.unwrap_or(false),
        show_r_squared: supports_equation && native_show_r_squared.unwrap_or(false),
    };
    value.validate()?;
    Ok(value)
}

fn patch_trendline(data: &[u8], expected: &ChartSeriesTrendline) -> Result<Vec<u8>> {
    expected.validate()?;
    let existing_extension = generated_chart_series_non_style_extension(data)?;
    if existing_extension.is_none() && expected == &ChartSeriesTrendline::none() {
        return Ok(data.to_vec());
    }
    let mut extension = existing_extension.unwrap_or_default().to_vec();
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension.as_slice())?;
    let visible = expected.trendline_type.is_visible();
    extension = patch_optional_varint(&extension, SHOW_TRENDLINE_FIELD, visible.then_some(1))?;
    extension = patch_optional_varint(
        &extension,
        TRENDLINE_TYPE_FIELD,
        visible.then_some(encode_i32(expected.trendline_type.native_type())),
    )?;

    if visible {
        extension = patch_optional_bytes(
            &extension,
            TRENDLINE_LABEL_FIELD,
            expected.custom_name.as_deref().map(str::as_bytes),
        )?;
        extension = patch_optional_varint(
            &extension,
            SHOW_LABEL_FIELD,
            expected.show_name_in_legend.then_some(1),
        )?;
        match expected.trendline_type {
            ChartSeriesTrendlineType::Polynomial(order) => {
                extension = patch_optional_varint(
                    &extension,
                    POLYNOMIAL_ORDER_FIELD,
                    Some(u64::from(order.get())),
                )?;
            },
            ChartSeriesTrendlineType::MovingAverage(period) => {
                extension = patch_optional_varint(
                    &extension,
                    MOVING_AVERAGE_PERIOD_FIELD,
                    Some(u64::from(period.get())),
                )?;
            },
            _ => {},
        }
        if expected.trendline_type.supports_equation() {
            extension = patch_optional_varint(
                &extension,
                SHOW_EQUATION_FIELD,
                expected.show_equation.then_some(1),
            )?;
            extension = patch_optional_varint(
                &extension,
                SHOW_R_SQUARED_FIELD,
                expected.show_r_squared.then_some(1),
            )?;
        }
    }

    let patched = patch_chart_series_non_style_extension(
        data,
        existing_extension.is_some(),
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    if read_trendline(&patched)? != *expected {
        return Err(Error::InvalidFormat(
            "chart series trendline wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

fn patch_optional_varint(data: &[u8], field_number: u32, value: Option<u64>) -> Result<Vec<u8>> {
    let present = strict_optional_varint(data, field_number)?.is_some();
    patch_varint_field(data, field_number, present, value)
}

fn patch_optional_bytes(data: &[u8], field_number: u32, value: Option<&[u8]>) -> Result<Vec<u8>> {
    let present = strict_optional_bytes(data, field_number)?.is_some();
    patch_length_delimited_field(data, field_number, present, value)
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

fn strict_optional_string(data: &[u8], field_number: u32) -> Result<Option<String>> {
    strict_optional_bytes(data, field_number)?
        .map(|bytes| {
            std::str::from_utf8(bytes)
                .map(str::to_owned)
                .map_err(|error| {
                    Error::InvalidFormat(format!(
                        "chart series trendline field {field_number} is not UTF-8: {error}"
                    ))
                })
        })
        .transpose()
}

fn strict_optional_bytes(data: &[u8], field_number: u32) -> Result<Option<&[u8]>> {
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
    if field.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart series trendline field {field_number} is not length-delimited"
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
            "singular chart series trendline field {field_number} occurs more than once"
        )));
    }
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart series trendline field {field_number} is not a varint"
        )));
    }
    let (value, consumed) =
        litchi_iwa_common::varint::decode_varint_from_bytes(&data[field.key_end..field.end])
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
    if litchi_iwa_common::varint::encoded_len(value) != consumed {
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
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_EXTENSION_FIELD: u32 = 4_097;

    #[test]
    fn every_native_inspector_type_round_trips_with_dependent_options() {
        let order = ChartSeriesTrendlinePolynomialOrder::new(5).unwrap();
        let period = ChartSeriesTrendlineMovingAveragePeriod::new(12).unwrap();
        let equation = ChartSeriesTrendline::linear()
            .with_legend_name("Linear fit")
            .unwrap()
            .with_equation_visibility(true)
            .unwrap()
            .with_r_squared_visibility(true)
            .unwrap();
        let variants = [
            equation,
            ChartSeriesTrendline::logarithmic(),
            ChartSeriesTrendline::polynomial(order),
            ChartSeriesTrendline::power(),
            ChartSeriesTrendline::exponential(),
            ChartSeriesTrendline::moving_average(period)
                .with_legend_visibility(true)
                .unwrap(),
        ];
        for trendline in variants {
            let original = canonical_empty_chart_series_non_style_data().unwrap();
            let patched = patch_trendline(&original, &trendline).unwrap();
            assert_eq!(read_trendline(&patched).unwrap(), trendline);
            let hidden = patch_trendline(&patched, &ChartSeriesTrendline::none()).unwrap();
            assert_eq!(
                read_trendline(&hidden).unwrap(),
                ChartSeriesTrendline::none()
            );
            assert_eq!(
                read_trendline(&patch_trendline(&hidden, &trendline).unwrap()).unwrap(),
                trendline
            );
        }
    }

    #[test]
    fn checked_parameters_and_display_dependencies_are_enforced() {
        assert!(ChartSeriesTrendlinePolynomialOrder::new(1).is_err());
        assert!(ChartSeriesTrendlinePolynomialOrder::new(7).is_err());
        assert!(ChartSeriesTrendlineMovingAveragePeriod::new(1).is_err());
        assert!(
            ChartSeriesTrendline::moving_average(ChartSeriesTrendlineMovingAveragePeriod::DEFAULT)
                .with_equation_visibility(true)
                .is_err()
        );
        assert!(
            ChartSeriesTrendline::none()
                .with_legend_visibility(true)
                .is_err()
        );
    }

    #[test]
    fn patch_preserves_neighboring_unknown_and_inactive_parameter_fields() {
        let mut extension = Vec::new();
        append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, 91).unwrap();
        append_varint_field(&mut extension, POLYNOMIAL_ORDER_FIELD, 6).unwrap();
        let mut original = canonical_empty_chart_series_non_style_data().unwrap();
        original =
            patch_chart_series_non_style_extension(&original, false, Some(extension.as_slice()))
                .unwrap();
        append_varint_field(&mut original, UNKNOWN_OUTER_FIELD, 73).unwrap();

        let expected = ChartSeriesTrendline::power()
            .with_legend_name("Power fit")
            .unwrap();
        let patched = patch_trendline(&original, &expected).unwrap();
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
        assert_eq!(
            strict_optional_i32(extension, POLYNOMIAL_ORDER_FIELD).unwrap(),
            Some(6)
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
        extension.extend(litchi_iwa_common::varint::encode_varint(
            u64::from(SHOW_TRENDLINE_FIELD) << 3,
        ));
        extension.extend([0x81, 0x00]);
        append_varint_field(&mut extension, TRENDLINE_TYPE_FIELD, NATIVE_LINEAR as u64).unwrap();
        let overlong_bool =
            patch_chart_series_non_style_extension(&original, false, Some(&extension)).unwrap();
        assert!(read_trendline(&overlong_bool).is_err());

        extension.clear();
        append_varint_field(&mut extension, SHOW_TRENDLINE_FIELD, 1).unwrap();
        append_varint_field(
            &mut extension,
            TRENDLINE_TYPE_FIELD,
            NATIVE_POLYNOMIAL as u64,
        )
        .unwrap();
        append_varint_field(&mut extension, POLYNOMIAL_ORDER_FIELD, 7).unwrap();
        let invalid_order =
            patch_chart_series_non_style_extension(&original, false, Some(&extension)).unwrap();
        assert!(read_trendline(&invalid_order).is_err());

        extension.clear();
        append_varint_field(&mut extension, SHOW_TRENDLINE_FIELD, 1).unwrap();
        append_varint_field(&mut extension, TRENDLINE_TYPE_FIELD, NATIVE_LINEAR as u64).unwrap();
        append_length_delimited_field(&mut extension, TRENDLINE_LABEL_FIELD, &[0xff]).unwrap();
        let invalid_name =
            patch_chart_series_non_style_extension(&original, false, Some(&extension)).unwrap();
        assert!(read_trendline(&invalid_name).is_err());

        extension.clear();
        append_varint_field(&mut extension, SHOW_TRENDLINE_FIELD, 1).unwrap();
        append_varint_field(
            &mut extension,
            TRENDLINE_TYPE_FIELD,
            NATIVE_MOVING_AVERAGE as u64,
        )
        .unwrap();
        append_varint_field(&mut extension, SHOW_EQUATION_FIELD, 2).unwrap();
        let malformed_dormant_option =
            patch_chart_series_non_style_extension(&original, false, Some(&extension)).unwrap();
        assert!(read_trendline(&malformed_dormant_option).is_err());
    }

    #[test]
    fn unsupported_values_are_lossless_but_known_aliases_are_rejected() {
        const FUTURE_TYPE: i32 = 9_001;
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let future = ChartSeriesTrendline::unsupported(FUTURE_TYPE).unwrap();
        let patched = patch_trendline(&original, &future).unwrap();
        assert_eq!(read_trendline(&patched).unwrap(), future);
        assert!(ChartSeriesTrendline::unsupported(NATIVE_LINEAR).is_err());
    }
}
