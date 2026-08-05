//! Scalar, enumeration, and range validation for chart XML attributes.
//!
//! Keeping these conversions in one layer makes semantic decoders explicit
//! about the object they are building and centralizes the format constraints.

use crate::chart::axis::{AxisCrossBetween, AxisCrossMode, AxisLabelAlign, BuiltInUnit, TimeUnit};
use crate::chart::data::NumberFormat;
use crate::chart::types::{
    AxisOrientation, AxisPosition, BarGrouping, DataLabelPosition, DisplayBlanks, MarkerStyle,
    TickLabelPosition, TickMark,
};
use crate::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;

#[inline]
pub(super) fn parse_grouping(e: &BytesStart<'_>) -> Result<BarGrouping> {
    if let Some(val) = get_attr(e, b"val") {
        Ok(match val.as_slice() {
            b"standard" => BarGrouping::Standard,
            b"clustered" => BarGrouping::Clustered,
            b"stacked" => BarGrouping::Stacked,
            b"percentStacked" => BarGrouping::PercentStacked,
            _ => return Err(invalid_attribute("chart grouping", &val)),
        })
    } else {
        Ok(BarGrouping::Standard)
    }
}

#[inline]
pub(super) fn parse_axis_position(e: &BytesStart<'_>) -> Result<AxisPosition> {
    if let Some(val) = get_attr(e, b"val") {
        Ok(match val.as_slice() {
            b"b" => AxisPosition::Bottom,
            b"l" => AxisPosition::Left,
            b"r" => AxisPosition::Right,
            b"t" => AxisPosition::Top,
            _ => return Err(invalid_attribute("chart axis position", &val)),
        })
    } else {
        Ok(AxisPosition::Bottom)
    }
}

pub(super) fn parse_axis_orientation(element: &BytesStart<'_>) -> Result<AxisOrientation> {
    let value =
        get_attr(element, b"val").ok_or_else(|| missing_attribute("chart axis orientation"))?;
    match value.as_slice() {
        b"minMax" => Ok(AxisOrientation::MinMax),
        b"maxMin" => Ok(AxisOrientation::MaxMin),
        _ => Err(invalid_attribute("chart axis orientation", &value)),
    }
}

pub(super) fn parse_tick_mark(element: &BytesStart<'_>) -> Result<TickMark> {
    let value =
        get_attr(element, b"val").ok_or_else(|| missing_attribute("chart tick-mark style"))?;
    match value.as_slice() {
        b"cross" => Ok(TickMark::Cross),
        b"in" => Ok(TickMark::In),
        b"none" => Ok(TickMark::None),
        b"out" => Ok(TickMark::Out),
        _ => Err(invalid_attribute("chart tick-mark style", &value)),
    }
}

pub(super) fn parse_tick_label_position(element: &BytesStart<'_>) -> Result<TickLabelPosition> {
    let value =
        get_attr(element, b"val").ok_or_else(|| missing_attribute("chart tick-label position"))?;
    match value.as_slice() {
        b"high" => Ok(TickLabelPosition::High),
        b"low" => Ok(TickLabelPosition::Low),
        b"nextTo" => Ok(TickLabelPosition::NextTo),
        b"none" => Ok(TickLabelPosition::None),
        _ => Err(invalid_attribute("chart tick-label position", &value)),
    }
}

pub(super) fn parse_axis_cross_mode(element: &BytesStart<'_>) -> Result<AxisCrossMode> {
    let value =
        get_attr(element, b"val").ok_or_else(|| missing_attribute("chart axis crossing mode"))?;
    match value.as_slice() {
        b"autoZero" => Ok(AxisCrossMode::AutoZero),
        b"max" => Ok(AxisCrossMode::Max),
        b"min" => Ok(AxisCrossMode::Min),
        _ => Err(invalid_attribute("chart axis crossing mode", &value)),
    }
}

pub(super) fn parse_axis_cross_between(element: &BytesStart<'_>) -> Result<AxisCrossBetween> {
    let value = get_attr(element, b"val")
        .ok_or_else(|| missing_attribute("chart axis crossing position"))?;
    match value.as_slice() {
        b"between" => Ok(AxisCrossBetween::Between),
        b"midCat" => Ok(AxisCrossBetween::MidCategory),
        _ => Err(invalid_attribute("chart axis crossing position", &value)),
    }
}

pub(super) fn parse_axis_label_align(element: &BytesStart<'_>) -> Result<AxisLabelAlign> {
    let value =
        get_attr(element, b"val").ok_or_else(|| missing_attribute("chart axis label alignment"))?;
    match value.as_slice() {
        b"ctr" => Ok(AxisLabelAlign::Center),
        b"l" => Ok(AxisLabelAlign::Left),
        b"r" => Ok(AxisLabelAlign::Right),
        _ => Err(invalid_attribute("chart axis label alignment", &value)),
    }
}

pub(super) fn parse_time_unit(element: &BytesStart<'_>) -> Result<TimeUnit> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute("chart time unit"))?;
    match value.as_slice() {
        b"days" => Ok(TimeUnit::Days),
        b"months" => Ok(TimeUnit::Months),
        b"years" => Ok(TimeUnit::Years),
        _ => Err(invalid_attribute("chart time unit", &value)),
    }
}

pub(super) fn parse_built_in_unit(element: &BytesStart<'_>) -> Result<BuiltInUnit> {
    let value = get_attr(element, b"val")
        .ok_or_else(|| missing_attribute("chart built-in display unit"))?;
    match value.as_slice() {
        b"hundreds" => Ok(BuiltInUnit::Hundreds),
        b"thousands" => Ok(BuiltInUnit::Thousands),
        b"tenThousands" => Ok(BuiltInUnit::TenThousands),
        b"hundredThousands" => Ok(BuiltInUnit::HundredThousands),
        b"millions" => Ok(BuiltInUnit::Millions),
        b"tenMillions" => Ok(BuiltInUnit::TenMillions),
        b"hundredMillions" => Ok(BuiltInUnit::HundredMillions),
        b"billions" => Ok(BuiltInUnit::Billions),
        b"trillions" => Ok(BuiltInUnit::Trillions),
        _ => Err(invalid_attribute("chart built-in display unit", &value)),
    }
}

pub(super) fn parse_marker_style(element: &BytesStart<'_>) -> Result<MarkerStyle> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute("chart marker style"))?;
    match value.as_slice() {
        b"circle" => Ok(MarkerStyle::Circle),
        b"dash" => Ok(MarkerStyle::Dash),
        b"diamond" => Ok(MarkerStyle::Diamond),
        b"dot" => Ok(MarkerStyle::Dot),
        b"none" => Ok(MarkerStyle::None),
        b"picture" => Ok(MarkerStyle::Picture),
        b"plus" => Ok(MarkerStyle::Plus),
        b"square" => Ok(MarkerStyle::Square),
        b"star" => Ok(MarkerStyle::Star),
        b"triangle" => Ok(MarkerStyle::Triangle),
        b"x" => Ok(MarkerStyle::X),
        b"auto" => Ok(MarkerStyle::Auto),
        _ => Err(invalid_attribute("chart marker style", &value)),
    }
}

pub(super) fn parse_data_label_position(element: &BytesStart<'_>) -> Result<DataLabelPosition> {
    let value =
        get_attr(element, b"val").ok_or_else(|| missing_attribute("chart data-label position"))?;
    match value.as_slice() {
        b"bestFit" => Ok(DataLabelPosition::BestFit),
        b"ctr" => Ok(DataLabelPosition::Center),
        b"inBase" => Ok(DataLabelPosition::InsideBase),
        b"inEnd" => Ok(DataLabelPosition::InsideEnd),
        b"l" => Ok(DataLabelPosition::Left),
        b"outEnd" => Ok(DataLabelPosition::OutsideEnd),
        b"r" => Ok(DataLabelPosition::Right),
        b"t" => Ok(DataLabelPosition::Top),
        b"b" => Ok(DataLabelPosition::Bottom),
        _ => Err(invalid_attribute("chart data-label position", &value)),
    }
}

pub(super) fn parse_number_format(
    element: &BytesStart<'_>,
    decoder: Decoder,
    description: &str,
) -> Result<NumberFormat> {
    let format_code = element
        .try_get_attribute(b"formatCode")
        .map_err(|error| Error::Xml(error.to_string()))?
        .ok_or_else(|| missing_attribute(&format!("{description} number format code")))?
        .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
        .map_err(|error| Error::Xml(error.to_string()))?
        .into_owned();
    let source_linked = match get_attr(element, b"sourceLinked") {
        Some(value) => parse_bool_value(&value, &format!("{description} source-linked flag"))?,
        None => true,
    };
    Ok(NumberFormat::new(format_code).with_source_linked(source_linked))
}

#[inline]
pub(super) fn parse_display_blanks(e: &BytesStart<'_>) -> Result<DisplayBlanks> {
    if let Some(val) = get_attr(e, b"val") {
        Ok(match val.as_slice() {
            b"gap" => DisplayBlanks::Gap,
            b"span" => DisplayBlanks::Span,
            b"zero" => DisplayBlanks::Zero,
            _ => return Err(invalid_attribute("chart blank-display mode", &val)),
        })
    } else {
        Ok(DisplayBlanks::Gap)
    }
}

#[inline]
pub(super) fn parse_bool_attr(e: &BytesStart<'_>) -> Result<bool> {
    if let Some(val) = get_attr(e, b"val") {
        parse_bool_value(&val, "chart boolean")
    } else {
        Ok(true)
    }
}

pub(super) fn parse_bool_value(value: &[u8], description: &str) -> Result<bool> {
    match value {
        b"1" | b"true" => Ok(true),
        b"0" | b"false" => Ok(false),
        _ => Err(invalid_attribute(description, value)),
    }
}

pub(super) fn invalid_attribute(description: &str, value: &[u8]) -> Error {
    Error::Invalid(format!(
        "invalid {description} '{}'",
        String::from_utf8_lossy(value)
    ))
}

pub(super) fn required_u32_attr(element: &BytesStart<'_>, description: &str) -> Result<u32> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute(description))?;
    std::str::from_utf8(&value)
        .ok()
        .and_then(|value| value.parse().ok())
        .ok_or_else(|| invalid_attribute(description, &value))
}

pub(super) fn required_string_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<String> {
    element
        .try_get_attribute(name)
        .map_err(|error| Error::Xml(error.to_string()))?
        .ok_or_else(|| missing_attribute(description))?
        .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
        .map(|value| value.into_owned())
        .map_err(|error| Error::Xml(error.to_string()))
}

pub(super) fn optional_u32_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    default: u32,
    description: &str,
) -> Result<u32> {
    let Some(value) = get_attr(element, name) else {
        return Ok(default);
    };
    std::str::from_utf8(&value)
        .ok()
        .and_then(|value| value.parse().ok())
        .ok_or_else(|| invalid_attribute(description, &value))
}

pub(super) fn optional_i32_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    default: i32,
    description: &str,
) -> Result<i32> {
    let Some(value) = get_attr(element, name) else {
        return Ok(default);
    };
    std::str::from_utf8(&value)
        .ok()
        .and_then(|value| value.parse().ok())
        .ok_or_else(|| invalid_attribute(description, &value))
}

pub(super) fn optional_bool_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    default: bool,
    description: &str,
) -> Result<bool> {
    match get_attr(element, name) {
        Some(value) => parse_bool_value(&value, description),
        None => Ok(default),
    }
}

pub(super) fn required_named_f64_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    description: &str,
) -> Result<f64> {
    let value = get_attr(element, name).ok_or_else(|| missing_attribute(description))?;
    std::str::from_utf8(&value)
        .ok()
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite())
        .ok_or_else(|| invalid_attribute(description, &value))
}

pub(super) fn required_positive_u32_attr(
    element: &BytesStart<'_>,
    description: &str,
) -> Result<u32> {
    let value = required_u32_attr(element, description)?;
    if value == 0 {
        return Err(Error::Invalid(format!("{description} must be positive")));
    }
    Ok(value)
}

pub(super) fn required_f64_attr(element: &BytesStart<'_>, description: &str) -> Result<f64> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute(description))?;
    let parsed = std::str::from_utf8(&value)
        .ok()
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite())
        .ok_or_else(|| invalid_attribute(description, &value))?;
    Ok(parsed)
}

pub(super) fn required_positive_f64_attr(
    element: &BytesStart<'_>,
    description: &str,
) -> Result<f64> {
    let value = required_f64_attr(element, description)?;
    if value <= 0.0 {
        return Err(Error::Invalid(format!("{description} must be positive")));
    }
    Ok(value)
}

pub(super) fn required_nonnegative_f64_attr(
    element: &BytesStart<'_>,
    description: &str,
) -> Result<f64> {
    let value = required_f64_attr(element, description)?;
    if value < 0.0 {
        return Err(Error::Invalid(format!("{description} must be nonnegative")));
    }
    Ok(value)
}

pub(super) fn required_enum_attr(element: &BytesStart<'_>, description: &str) -> Result<String> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute(description))?;
    String::from_utf8(value).map_err(|error| Error::Invalid(error.to_string()))
}

pub(super) fn bounded_u32_attr(
    element: &BytesStart<'_>,
    description: &str,
    minimum: u32,
    maximum: u32,
) -> Result<u32> {
    let value = required_u32_attr(element, description)?;
    if !(minimum..=maximum).contains(&value) {
        return Err(Error::Invalid(format!(
            "{description} must be between {minimum} and {maximum}"
        )));
    }
    Ok(value)
}

pub(super) fn bounded_percentage_u32_attr(
    element: &BytesStart<'_>,
    description: &str,
    minimum: u32,
    maximum: u32,
) -> Result<u32> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute(description))?;
    let text = std::str::from_utf8(&value).map_err(|_| invalid_attribute(description, &value))?;
    let digits = text.strip_suffix('%').unwrap_or(text);
    let parsed = digits
        .parse::<u32>()
        .map_err(|_| invalid_attribute(description, &value))?;
    if !(minimum..=maximum).contains(&parsed) {
        return Err(Error::Invalid(format!(
            "{description} must be between {minimum} and {maximum}"
        )));
    }
    Ok(parsed)
}

pub(super) fn bounded_percentage_i32_attr(
    element: &BytesStart<'_>,
    description: &str,
    minimum: i32,
    maximum: i32,
) -> Result<i32> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute(description))?;
    let text = std::str::from_utf8(&value).map_err(|_| invalid_attribute(description, &value))?;
    let digits = text.strip_suffix('%').unwrap_or(text);
    let parsed = digits
        .parse::<i32>()
        .map_err(|_| invalid_attribute(description, &value))?;
    if !(minimum..=maximum).contains(&parsed) {
        return Err(Error::Invalid(format!(
            "{description} must be between {minimum} and {maximum}"
        )));
    }
    Ok(parsed)
}

pub(super) fn missing_attribute(description: &str) -> Error {
    Error::Invalid(format!("{description} is missing its value"))
}

#[inline]
pub(super) fn get_attr(e: &BytesStart<'_>, name: &[u8]) -> Option<Vec<u8>> {
    e.attributes()
        .filter_map(|a| a.ok())
        .find(|a| a.key.as_ref() == name)
        .map(|a| a.value.to_vec())
}
