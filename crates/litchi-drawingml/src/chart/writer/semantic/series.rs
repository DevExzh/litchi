//! Series, data-point, data-label, trendline, and error-bar records.

use super::super::model::SeriesFeatures;
use super::super::validation::invalid_chart_input;
use super::super::xml::write_fragment;
use super::{
    common::write_chart_lines,
    presentation::{write_layout, write_marker_parts, write_picture_options, write_title_text},
};
use crate::chart::data::{NumericData, StringData, TitleText};
use crate::chart::series::{
    DataLabel, DataLabels, DataPoint, ErrorBar, ErrorBarDirection, ErrorBarType, ErrorBarValueType,
    Series, Trendline, TrendlineType,
};
use litchi_core::xml::escape_xml;
use std::io::Write;

pub(super) fn write_series<W: Write>(
    writer: &mut W,
    series: &Series,
    features: SeriesFeatures,
) -> std::io::Result<()> {
    write!(writer, "<c:ser>")?;
    write!(writer, r#"<c:idx val="{}"/>"#, series.index)?;
    write!(writer, r#"<c:order val="{}"/>"#, series.order)?;

    if let Some(title) = &series.title {
        write!(writer, "<c:tx>")?;
        match title {
            TitleText::Literal(rich_text) => {
                write!(writer, "<c:v>{}</c:v>", escape_xml(&rich_text.text))?;
            },
            TitleText::Reference(source_ref) => {
                write!(writer, "<c:strRef>")?;
                write!(writer, "<c:f>{}</c:f>", escape_xml(&source_ref.formula))?;
                write!(writer, "</c:strRef>")?;
            },
        }
        write!(writer, "</c:tx>")?;
    }

    if let Some(shape_properties) = series.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }

    write_series_presentation(writer, series, features)?;

    if let Some(ref categories) = series.categories {
        write_string_data_ref(writer, "c:cat", categories)?;
    }

    if let Some(ref values) = series.values {
        write_numeric_data_ref(writer, "c:val", values)?;
    }

    if !features.bar_shape && series.bar_shape.is_some() {
        return Err(invalid_chart_input(
            "chart type does not support per-series bar shapes",
        ));
    }
    if features.bar_shape
        && let Some(shape) = series.bar_shape
    {
        write!(writer, r#"<c:shape val="{}"/>"#, shape.xml_value())?;
    }

    if features.smooth {
        write!(
            writer,
            r#"<c:smooth val="{}"/>"#,
            if series.smooth { "1" } else { "0" }
        )?;
    }

    if let Some(extension_list) = series.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }

    write!(writer, "</c:ser>")?;

    Ok(())
}

pub(super) fn write_series_presentation<W: Write>(
    writer: &mut W,
    series: &Series,
    features: SeriesFeatures,
) -> std::io::Result<()> {
    if !features.marker
        && (series.marker_symbol.is_some()
            || series.marker_size.is_some()
            || series.marker_present
            || series.marker_shape_properties.is_some()
            || series.marker_extension_list.is_some())
    {
        return Err(invalid_chart_input(
            "chart type does not support series markers",
        ));
    }
    if !features.smooth && series.smooth {
        return Err(invalid_chart_input(
            "chart type does not support smoothed series",
        ));
    }
    if !features.invert_if_negative && series.invert_if_negative {
        return Err(invalid_chart_input(
            "chart type does not support negative-value inversion",
        ));
    }
    if !features.picture_options && series.picture_options.is_some() {
        return Err(invalid_chart_input(
            "chart type does not support series picture options",
        ));
    }
    if !features.explosion && series.explosion.is_some() {
        return Err(invalid_chart_input(
            "chart type does not support series explosion",
        ));
    }
    if !features.point_and_label_overrides
        && (!series.data_points.is_empty() || series.data_labels.is_some())
    {
        return Err(invalid_chart_input(
            "chart type does not support point or data-label overrides",
        ));
    }
    if !features.error_bars && !series.error_bars.is_empty() {
        return Err(invalid_chart_input(
            "chart type does not support error bars",
        ));
    }
    if !features.trendlines && !series.trendlines.is_empty() {
        return Err(invalid_chart_input(
            "chart type does not support trendlines",
        ));
    }
    if series
        .marker_size
        .is_some_and(|size| !(2..=72).contains(&size))
    {
        return Err(invalid_chart_input("chart series marker size must be 2-72"));
    }
    if features.marker
        && (series.marker_symbol.is_some()
            || series.marker_size.is_some()
            || series.marker_present
            || series.marker_shape_properties.is_some()
            || series.marker_extension_list.is_some())
    {
        write_marker_parts(
            writer,
            series.marker_symbol,
            series.marker_size,
            series.marker_shape_properties.as_ref(),
            series.marker_extension_list.as_ref(),
            "chart series",
        )?;
    }

    if features.invert_if_negative && series.invert_if_negative {
        write!(writer, r#"<c:invertIfNegative val="1"/>"#)?;
    }
    if features.picture_options
        && let Some(options) = series.picture_options.as_ref()
    {
        write_picture_options(writer, options)?;
    }
    if features.explosion
        && let Some(explosion) = series.explosion
    {
        write!(writer, r#"<c:explosion val="{}"/>"#, explosion)?;
    }
    for (position, point) in series.data_points.iter().enumerate() {
        if series.data_points[..position]
            .iter()
            .any(|existing| existing.index == point.index)
        {
            return Err(invalid_chart_input(format!(
                "duplicate chart data-point index {}",
                point.index
            )));
        }
        write_data_point(writer, point)?;
    }
    if let Some(labels) = &series.data_labels {
        write_data_labels(writer, labels)?;
    }
    for trendline in &series.trendlines {
        write_trendline(writer, trendline)?;
    }
    for error_bar in &series.error_bars {
        write_error_bar(writer, error_bar)?;
    }
    Ok(())
}

pub(super) fn write_data_point<W: Write>(writer: &mut W, point: &DataPoint) -> std::io::Result<()> {
    if point
        .marker_size
        .is_some_and(|size| !(2..=72).contains(&size))
    {
        return Err(invalid_chart_input(
            "chart data-point marker size must be 2-72",
        ));
    }
    write!(writer, r#"<c:dPt><c:idx val="{}"/>"#, point.index)?;
    if point.invert_if_negative {
        write!(writer, r#"<c:invertIfNegative val="1"/>"#)?;
    }
    if point.marker_present
        || point.marker_symbol.is_some()
        || point.marker_size.is_some()
        || point.marker_shape_properties.is_some()
        || point.marker_extension_list.is_some()
    {
        write_marker_parts(
            writer,
            point.marker_symbol,
            point.marker_size,
            point.marker_shape_properties.as_ref(),
            point.marker_extension_list.as_ref(),
            "chart data-point",
        )?;
    }
    if let Some(bubble_3d) = point.bubble_3d {
        write!(
            writer,
            r#"<c:bubble3D val="{}"/>"#,
            if bubble_3d { "1" } else { "0" }
        )?;
    }
    if let Some(explosion) = point.explosion {
        write!(writer, r#"<c:explosion val="{}"/>"#, explosion)?;
    }
    if let Some(shape_properties) = point.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    if let Some(options) = point.picture_options.as_ref() {
        write_picture_options(writer, options)?;
    }
    if let Some(extension_list) = point.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }
    write!(writer, "</c:dPt>")?;
    Ok(())
}

pub(super) fn write_data_labels<W: Write>(
    writer: &mut W,
    labels: &DataLabels,
) -> std::io::Result<()> {
    write!(writer, "<c:dLbls>")?;
    let mut point_indexes = std::collections::HashSet::with_capacity(labels.labels.len());
    for label in &labels.labels {
        if !point_indexes.insert(label.index) {
            return Err(invalid_chart_input(format!(
                "chart data labels contain duplicate point index {}",
                label.index
            )));
        }
        write_data_label(writer, label)?;
    }
    if labels.deleted {
        if labels.number_format.is_some()
            || labels.shape_properties.is_some()
            || labels.text_properties.is_some()
            || labels.position.is_some()
            || labels.show_legend_key
            || labels.show_value
            || labels.show_category_name
            || labels.show_series_name
            || labels.show_percent
            || labels.show_bubble_size
            || labels.separator.is_some()
            || labels.show_leader_lines
            || labels.leader_lines.is_some()
        {
            return Err(invalid_chart_input(
                "chart data labels cannot mix deletion with shared settings",
            ));
        }
        write!(writer, r#"<c:delete val="1"/>"#)?;
        if let Some(extension_list) = labels.extension_list.as_ref() {
            write_fragment(writer, extension_list.as_xml())?;
        }
        write!(writer, "</c:dLbls>")?;
        return Ok(());
    }
    if let Some(number_format) = labels.number_format.as_ref() {
        write!(
            writer,
            r#"<c:numFmt formatCode="{}" sourceLinked="{}"/>"#,
            escape_xml(&number_format.format_code),
            if number_format.source_linked {
                "1"
            } else {
                "0"
            }
        )?;
    }
    if let Some(shape_properties) = labels.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    if let Some(text_properties) = labels.text_properties.as_ref() {
        write_fragment(writer, text_properties.as_xml())?;
    }
    if let Some(position) = labels.position {
        write!(writer, r#"<c:dLblPos val="{}"/>"#, position.xml_value())?;
    }
    for (name, value) in [
        ("showLegendKey", labels.show_legend_key),
        ("showVal", labels.show_value),
        ("showCatName", labels.show_category_name),
        ("showSerName", labels.show_series_name),
        ("showPercent", labels.show_percent),
        ("showBubbleSize", labels.show_bubble_size),
    ] {
        write!(
            writer,
            r#"<c:{name} val="{}"/>"#,
            if value { "1" } else { "0" }
        )?;
    }
    if let Some(separator) = &labels.separator {
        write!(
            writer,
            "<c:separator>{}</c:separator>",
            escape_xml(separator)
        )?;
    }
    if labels.show_leader_lines {
        write!(writer, r#"<c:showLeaderLines val="1"/>"#)?;
    }
    if let Some(lines) = labels.leader_lines.as_ref() {
        write_chart_lines(writer, "leaderLines", lines)?;
    }
    if let Some(extension_list) = labels.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }
    write!(writer, "</c:dLbls>")?;
    Ok(())
}

pub(super) fn write_data_label<W: Write>(writer: &mut W, label: &DataLabel) -> std::io::Result<()> {
    write!(writer, r#"<c:dLbl><c:idx val="{}"/>"#, label.index)?;
    if label.deleted {
        if label.layout.is_some()
            || label.text.is_some()
            || label.number_format.is_some()
            || label.shape_properties.is_some()
            || label.text_properties.is_some()
            || label.position.is_some()
            || label.show_legend_key
            || label.show_value
            || label.show_category_name
            || label.show_series_name
            || label.show_percent
            || label.show_bubble_size
            || label.separator.is_some()
        {
            return Err(invalid_chart_input(
                "chart point data label cannot mix deletion with label settings",
            ));
        }
        write!(writer, r#"<c:delete val="1"/>"#)?;
        if let Some(extension_list) = label.extension_list.as_ref() {
            write_fragment(writer, extension_list.as_xml())?;
        }
        write!(writer, "</c:dLbl>")?;
        return Ok(());
    }
    if let Some(layout) = label.layout.as_ref() {
        write_layout(writer, Some(layout))?;
    }
    if let Some(text) = label.text.as_ref() {
        write_title_text(writer, text)?;
    }
    if let Some(number_format) = label.number_format.as_ref() {
        write!(
            writer,
            r#"<c:numFmt formatCode="{}" sourceLinked="{}"/>"#,
            escape_xml(&number_format.format_code),
            if number_format.source_linked {
                "1"
            } else {
                "0"
            }
        )?;
    }
    if let Some(shape_properties) = label.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    if let Some(text_properties) = label.text_properties.as_ref() {
        write_fragment(writer, text_properties.as_xml())?;
    }
    if let Some(position) = label.position {
        write!(writer, r#"<c:dLblPos val="{}"/>"#, position.xml_value())?;
    }
    for (name, value) in [
        ("showLegendKey", label.show_legend_key),
        ("showVal", label.show_value),
        ("showCatName", label.show_category_name),
        ("showSerName", label.show_series_name),
        ("showPercent", label.show_percent),
        ("showBubbleSize", label.show_bubble_size),
    ] {
        write!(
            writer,
            r#"<c:{name} val="{}"/>"#,
            if value { "1" } else { "0" }
        )?;
    }
    if let Some(separator) = label.separator.as_ref() {
        write!(
            writer,
            "<c:separator>{}</c:separator>",
            escape_xml(separator)
        )?;
    }
    if let Some(extension_list) = label.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }
    write!(writer, "</c:dLbl>")?;
    Ok(())
}

pub(super) fn write_trendline<W: Write>(
    writer: &mut W,
    trendline: &Trendline,
) -> std::io::Result<()> {
    validate_trendline(trendline)?;
    write!(writer, "<c:trendline>")?;
    if let Some(name) = &trendline.name {
        write!(writer, "<c:name>{}</c:name>", escape_xml(name))?;
    }
    if let Some(shape_properties) = trendline.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    let kind = match trendline.trendline_type {
        TrendlineType::Exponential => "exp",
        TrendlineType::Linear => "linear",
        TrendlineType::Logarithmic => "log",
        TrendlineType::MovingAverage => "movingAvg",
        TrendlineType::Polynomial => "poly",
        TrendlineType::Power => "power",
    };
    write!(writer, r#"<c:trendlineType val="{kind}"/>"#)?;
    for (name, value) in [("order", trendline.order), ("period", trendline.period)] {
        if let Some(value) = value {
            write!(writer, r#"<c:{name} val="{value}"/>"#)?;
        }
    }
    for (name, value) in [
        ("forward", trendline.forward),
        ("backward", trendline.backward),
        ("intercept", trendline.intercept),
    ] {
        if let Some(value) = value {
            write!(writer, r#"<c:{name} val="{value}"/>"#)?;
        }
    }
    write!(
        writer,
        r#"<c:dispRSqr val="{}"/><c:dispEq val="{}"/>"#,
        if trendline.display_r_squared {
            "1"
        } else {
            "0"
        },
        if trendline.display_equation { "1" } else { "0" }
    )?;
    if trendline.show_label
        || trendline.label.is_some()
        || trendline.label_layout.is_some()
        || trendline.label_number_format.is_some()
        || trendline.label_shape_properties.is_some()
        || trendline.label_text_properties.is_some()
        || trendline.label_extension_list.is_some()
    {
        write!(writer, "<c:trendlineLbl>")?;
        if let Some(layout) = trendline.label_layout.as_ref() {
            write_layout(writer, Some(layout))?;
        }
        if let Some(label) = trendline.label.as_ref() {
            write_title_text(writer, label)?;
        }
        if let Some(number_format) = trendline.label_number_format.as_ref() {
            write!(
                writer,
                r#"<c:numFmt formatCode="{}" sourceLinked="{}"/>"#,
                escape_xml(&number_format.format_code),
                if number_format.source_linked {
                    "1"
                } else {
                    "0"
                }
            )?;
        }
        if let Some(shape_properties) = trendline.label_shape_properties.as_ref() {
            write_fragment(writer, shape_properties.as_xml())?;
        }
        if let Some(text_properties) = trendline.label_text_properties.as_ref() {
            write_fragment(writer, text_properties.as_xml())?;
        }
        if let Some(extension_list) = trendline.label_extension_list.as_ref() {
            write_fragment(writer, extension_list.as_xml())?;
        }
        write!(writer, "</c:trendlineLbl>")?;
    }
    if let Some(extension_list) = trendline.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }
    write!(writer, "</c:trendline>")?;
    Ok(())
}

pub(super) fn validate_trendline(trendline: &Trendline) -> std::io::Result<()> {
    match trendline.trendline_type {
        TrendlineType::Polynomial if !matches!(trendline.order, Some(2..=6)) => {
            return Err(invalid_chart_input(
                "polynomial trendline order must be 2-6",
            ));
        },
        TrendlineType::MovingAverage if !matches!(trendline.period, Some(2..=255)) => {
            return Err(invalid_chart_input(
                "moving-average trendline period must be 2-255",
            ));
        },
        _ => {},
    }
    if !matches!(trendline.trendline_type, TrendlineType::Polynomial) && trendline.order.is_some() {
        return Err(invalid_chart_input(
            "only polynomial trendlines can specify an order",
        ));
    }
    if !matches!(trendline.trendline_type, TrendlineType::MovingAverage)
        && trendline.period.is_some()
    {
        return Err(invalid_chart_input(
            "only moving-average trendlines can specify a period",
        ));
    }
    for (name, value) in [
        ("forward", trendline.forward),
        ("backward", trendline.backward),
    ] {
        if value.is_some_and(|value| !value.is_finite() || value < 0.0) {
            return Err(invalid_chart_input(format!(
                "trendline {name} value must be finite and nonnegative"
            )));
        }
    }
    if trendline.intercept.is_some_and(|value| !value.is_finite()) {
        return Err(invalid_chart_input("trendline intercept must be finite"));
    }
    Ok(())
}

pub(super) fn write_error_bar<W: Write>(
    writer: &mut W,
    error_bar: &ErrorBar,
) -> std::io::Result<()> {
    validate_error_bar(error_bar)?;
    let direction = match error_bar.direction {
        ErrorBarDirection::X => "x",
        ErrorBarDirection::Y => "y",
    };
    let bar_type = match error_bar.error_type {
        ErrorBarType::Both => "both",
        ErrorBarType::Plus => "plus",
        ErrorBarType::Minus => "minus",
    };
    let value_type = match error_bar.value_type {
        ErrorBarValueType::Fixed => "fixedVal",
        ErrorBarValueType::Percentage => "percentage",
        ErrorBarValueType::StdDev => "stdDev",
        ErrorBarValueType::StdErr => "stdErr",
        ErrorBarValueType::Custom => "cust",
    };
    write!(
        writer,
        r#"<c:errBars><c:errDir val="{direction}"/><c:errBarType val="{bar_type}"/><c:errValType val="{value_type}"/>"#
    )?;
    write!(
        writer,
        r#"<c:noEndCap val="{}"/>"#,
        if error_bar.no_end_cap { "1" } else { "0" }
    )?;
    if let Some(values) = &error_bar.plus_values {
        write_numeric_data_ref(writer, "c:plus", values)?;
    }
    if let Some(values) = &error_bar.minus_values {
        write_numeric_data_ref(writer, "c:minus", values)?;
    }
    if let Some(value) = error_bar.value {
        write!(writer, r#"<c:val val="{}"/>"#, value)?;
    }
    if let Some(shape_properties) = error_bar.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    if let Some(extension_list) = error_bar.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }
    write!(writer, "</c:errBars>")?;
    Ok(())
}

pub(super) fn validate_error_bar(error_bar: &ErrorBar) -> std::io::Result<()> {
    if error_bar
        .value
        .is_some_and(|value| !value.is_finite() || value < 0.0)
    {
        return Err(invalid_chart_input(
            "error-bar value must be finite and nonnegative",
        ));
    }
    match error_bar.value_type {
        ErrorBarValueType::Fixed | ErrorBarValueType::Percentage | ErrorBarValueType::StdDev
            if error_bar.value.is_none() =>
        {
            Err(invalid_chart_input(
                "fixed, percentage, and standard-deviation error bars require a value",
            ))
        },
        ErrorBarValueType::Custom
            if error_bar.plus_values.is_none() && error_bar.minus_values.is_none() =>
        {
            Err(invalid_chart_input(
                "custom error bars require plus or minus values",
            ))
        },
        ErrorBarValueType::StdErr | ErrorBarValueType::Custom if error_bar.value.is_some() => Err(
            invalid_chart_input("standard-error and custom error bars cannot have a scalar value"),
        ),
        _ => Ok(()),
    }
}

pub(super) fn write_scatter_series<W: Write>(
    writer: &mut W,
    series: &Series,
) -> std::io::Result<()> {
    if series.bubble_sizes.is_some() {
        return Err(invalid_chart_input(
            "scatter series cannot contain bubble sizes",
        ));
    }
    write!(writer, "<c:ser>")?;
    write!(writer, r#"<c:idx val="{}"/>"#, series.index)?;
    write!(writer, r#"<c:order val="{}"/>"#, series.order)?;

    if let Some(title) = &series.title {
        write!(writer, "<c:tx>")?;
        match title {
            TitleText::Literal(rich_text) => {
                write!(writer, "<c:v>{}</c:v>", escape_xml(&rich_text.text))?;
            },
            TitleText::Reference(source_ref) => {
                write!(writer, "<c:strRef>")?;
                write!(writer, "<c:f>{}</c:f>", escape_xml(&source_ref.formula))?;
                write!(writer, "</c:strRef>")?;
            },
        }
        write!(writer, "</c:tx>")?;
    }

    if let Some(shape_properties) = series.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }

    write_series_presentation(writer, series, SeriesFeatures::LINE)?;

    if let Some(ref x_values) = series.x_values {
        write_numeric_data_ref(writer, "c:xVal", x_values)?;
    }

    if let Some(ref y_values) = series.y_values {
        write_numeric_data_ref(writer, "c:yVal", y_values)?;
    }

    write!(
        writer,
        r#"<c:smooth val="{}"/>"#,
        if series.smooth { "1" } else { "0" }
    )?;

    if let Some(extension_list) = series.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }

    write!(writer, "</c:ser>")?;

    Ok(())
}

pub(super) fn write_bubble_series<W: Write>(
    writer: &mut W,
    series: &Series,
) -> std::io::Result<()> {
    write!(writer, "<c:ser>")?;
    write!(writer, r#"<c:idx val="{}"/>"#, series.index)?;
    write!(writer, r#"<c:order val="{}"/>"#, series.order)?;

    if let Some(title) = &series.title {
        write!(writer, "<c:tx>")?;
        match title {
            TitleText::Literal(rich_text) => {
                write!(writer, "<c:v>{}</c:v>", escape_xml(&rich_text.text))?;
            },
            TitleText::Reference(source_ref) => {
                write!(writer, "<c:strRef>")?;
                write!(writer, "<c:f>{}</c:f>", escape_xml(&source_ref.formula))?;
                write!(writer, "</c:strRef>")?;
            },
        }
        write!(writer, "</c:tx>")?;
    }

    if let Some(shape_properties) = series.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }

    write_series_presentation(writer, series, SeriesFeatures::BUBBLE)?;

    if let Some(ref x_values) = series.x_values {
        write_numeric_data_ref(writer, "c:xVal", x_values)?;
    }

    if let Some(ref y_values) = series.y_values {
        write_numeric_data_ref(writer, "c:yVal", y_values)?;
    }

    if let Some(ref bubble_sizes) = series.bubble_sizes {
        write_numeric_data_ref(writer, "c:bubbleSize", bubble_sizes)?;
    }

    write!(
        writer,
        r#"<c:bubble3D val="{}"/>"#,
        if series.bubble_3d { "1" } else { "0" }
    )?;

    if let Some(extension_list) = series.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }

    write!(writer, "</c:ser>")?;

    Ok(())
}

pub(super) fn write_string_data_ref<W: Write>(
    writer: &mut W,
    tag: &str,
    data: &StringData,
) -> std::io::Result<()> {
    write!(writer, "<{}>", tag)?;

    if let Some(ref source_ref) = data.source_ref {
        write!(writer, "<c:strRef>")?;
        write!(writer, "<c:f>{}</c:f>", escape_xml(&source_ref.formula))?;

        if !data.values.is_empty() {
            write!(writer, "<c:strCache>")?;
            write!(writer, r#"<c:ptCount val="{}"/>"#, data.values.len())?;
            for (i, val) in data.values.iter().enumerate() {
                write!(
                    writer,
                    r#"<c:pt idx="{}"><c:v>{}</c:v></c:pt>"#,
                    i,
                    escape_xml(val)
                )?;
            }
            write!(writer, "</c:strCache>")?;
        }

        write!(writer, "</c:strRef>")?;
    } else if !data.values.is_empty() {
        write!(writer, "<c:strLit>")?;
        write!(writer, r#"<c:ptCount val="{}"/>"#, data.values.len())?;
        for (i, val) in data.values.iter().enumerate() {
            write!(
                writer,
                r#"<c:pt idx="{}"><c:v>{}</c:v></c:pt>"#,
                i,
                escape_xml(val)
            )?;
        }
        write!(writer, "</c:strLit>")?;
    }

    write!(writer, "</{}>", tag)?;

    Ok(())
}

pub(super) fn write_numeric_data_ref<W: Write>(
    writer: &mut W,
    tag: &str,
    data: &NumericData,
) -> std::io::Result<()> {
    write!(writer, "<{}>", tag)?;

    if let Some(ref source_ref) = data.source_ref {
        write!(writer, "<c:numRef>")?;
        write!(writer, "<c:f>{}</c:f>", escape_xml(&source_ref.formula))?;

        if !data.values.is_empty() {
            write!(writer, "<c:numCache>")?;
            write!(
                writer,
                r#"<c:formatCode>{}</c:formatCode>"#,
                escape_xml(data.format_code.as_deref().unwrap_or("General"))
            )?;
            write!(writer, r#"<c:ptCount val="{}"/>"#, data.values.len())?;
            for (i, val) in data.values.iter().enumerate() {
                write!(writer, r#"<c:pt idx="{}"><c:v>{}</c:v></c:pt>"#, i, val)?;
            }
            write!(writer, "</c:numCache>")?;
        }

        write!(writer, "</c:numRef>")?;
    } else if !data.values.is_empty() {
        write!(writer, "<c:numLit>")?;
        write!(
            writer,
            r#"<c:formatCode>{}</c:formatCode>"#,
            escape_xml(data.format_code.as_deref().unwrap_or("General"))
        )?;
        write!(writer, r#"<c:ptCount val="{}"/>"#, data.values.len())?;
        for (i, val) in data.values.iter().enumerate() {
            write!(writer, r#"<c:pt idx="{}"><c:v>{}</c:v></c:pt>"#, i, val)?;
        }
        write!(writer, "</c:numLit>")?;
    }

    write!(writer, "</{}>", tag)?;

    Ok(())
}
