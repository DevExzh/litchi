//! Axis record families and shared axis serialization.

use super::super::validation::invalid_chart_input;
use super::super::xml::write_fragment;
use super::{
    common::write_chart_lines,
    presentation::{write_layout, write_title, write_title_text},
};
use crate::chart::axis::{Axis, AxisCommon, CategoryAxis, DateAxis, SeriesAxis, ValueAxis};
use litchi_core::xml::escape_xml;
use std::io::Write;

pub(super) fn write_axis<W: Write>(writer: &mut W, axis: &Axis) -> std::io::Result<()> {
    match axis {
        Axis::Category(ax) => write_category_axis(writer, ax),
        Axis::Value(ax) => write_value_axis(writer, ax),
        Axis::Date(ax) => write_date_axis(writer, ax),
        Axis::Series(ax) => write_series_axis(writer, ax),
    }
}

pub(super) fn write_axis_common<W: Write>(
    writer: &mut W,
    common: &AxisCommon,
    min: Option<f64>,
    max: Option<f64>,
    log_base: Option<f64>,
) -> std::io::Result<()> {
    if min.is_some_and(|value| !value.is_finite()) || max.is_some_and(|value| !value.is_finite()) {
        return Err(invalid_chart_input("chart axis bounds must be finite"));
    }
    if min.zip(max).is_some_and(|(min, max)| min > max) {
        return Err(invalid_chart_input(
            "chart axis minimum cannot exceed its maximum",
        ));
    }
    if log_base.is_some_and(|base| !base.is_finite() || !(2.0..=1000.0).contains(&base)) {
        return Err(invalid_chart_input(
            "chart logarithmic base must be between 2 and 1000",
        ));
    }
    write!(writer, r#"<c:axId val="{}"/>"#, common.axis_id)?;

    write!(writer, "<c:scaling>")?;
    if let Some(log_base) = log_base {
        write!(writer, r#"<c:logBase val="{}"/>"#, log_base)?;
    }
    write!(
        writer,
        r#"<c:orientation val="{}"/>"#,
        common.orientation.xml_value()
    )?;
    if let Some(max) = max {
        write!(writer, r#"<c:max val="{}"/>"#, max)?;
    }
    if let Some(min) = min {
        write!(writer, r#"<c:min val="{}"/>"#, min)?;
    }
    if let Some(extension_list) = common.scaling_extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }
    write!(writer, "</c:scaling>")?;

    write!(
        writer,
        r#"<c:delete val="{}"/>"#,
        if common.deleted { "1" } else { "0" }
    )?;

    write!(
        writer,
        r#"<c:axPos val="{}"/>"#,
        common.position.xml_value()
    )?;

    if let Some(lines) = common.major_gridlines.as_ref() {
        write_chart_lines(writer, "majorGridlines", lines)?;
    } else if common.show_major_gridlines {
        write!(writer, "<c:majorGridlines/>")?;
    }

    if let Some(lines) = common.minor_gridlines.as_ref() {
        write_chart_lines(writer, "minorGridlines", lines)?;
    } else if common.show_minor_gridlines {
        write!(writer, "<c:minorGridlines/>")?;
    }

    if let Some(ref title) = common.title {
        write_title(
            writer,
            title,
            common.layout.as_ref(),
            common.title_overlay,
            common.title_shape_properties.as_ref(),
            common.title_text_properties.as_ref(),
            common.title_extension_list.as_ref(),
        )?;
    }

    if let Some(number_format) = &common.number_format {
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

    write!(
        writer,
        r#"<c:majorTickMark val="{}"/>"#,
        common.major_tick_mark.xml_value()
    )?;
    write!(
        writer,
        r#"<c:minorTickMark val="{}"/>"#,
        common.minor_tick_mark.xml_value()
    )?;
    write!(
        writer,
        r#"<c:tickLblPos val="{}"/>"#,
        common.tick_label_position.xml_value()
    )?;

    if let Some(shape_properties) = common.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    if let Some(text_properties) = common.text_properties.as_ref() {
        write_fragment(writer, text_properties.as_xml())?;
    }

    write!(writer, r#"<c:crossAx val="{}"/>"#, common.cross_axis_id)?;

    if let Some(crosses_at) = common.crosses_at {
        write!(writer, r#"<c:crossesAt val="{}"/>"#, crosses_at)?;
    } else {
        write!(
            writer,
            r#"<c:crosses val="{}"/>"#,
            common.cross_mode.xml_value()
        )?;
    }

    Ok(())
}

pub(super) fn write_category_axis<W: Write>(
    writer: &mut W,
    axis: &CategoryAxis,
) -> std::io::Result<()> {
    write!(writer, "<c:catAx>")?;
    write_axis_common(writer, &axis.common, axis.min, axis.max, axis.log_base)?;
    write!(
        writer,
        r#"<c:auto val="{}"/>"#,
        if axis.auto { "1" } else { "0" }
    )?;
    write!(
        writer,
        r#"<c:lblAlgn val="{}"/>"#,
        axis.label_align.map(|a| a.xml_value()).unwrap_or("ctr")
    )?;
    write!(
        writer,
        r#"<c:lblOffset val="{}"/>"#,
        axis.label_offset.unwrap_or(100)
    )?;
    if let Some(skip) = axis.tick_label_skip {
        write!(writer, r#"<c:tickLblSkip val="{}"/>"#, skip)?;
    }
    if let Some(skip) = axis.tick_mark_skip {
        write!(writer, r#"<c:tickMarkSkip val="{}"/>"#, skip)?;
    }
    write!(
        writer,
        r#"<c:noMultiLvlLbl val="{}"/>"#,
        if axis.no_multi_level { "1" } else { "0" }
    )?;
    write_axis_extension(writer, &axis.common)?;
    write!(writer, "</c:catAx>")?;
    Ok(())
}

pub(super) fn write_value_axis<W: Write>(writer: &mut W, axis: &ValueAxis) -> std::io::Result<()> {
    write!(writer, "<c:valAx>")?;
    write_axis_common(writer, &axis.common, axis.min, axis.max, axis.log_base)?;

    write!(
        writer,
        r#"<c:crossBetween val="{}"/>"#,
        axis.cross_between.xml_value()
    )?;

    if let Some(major_unit) = axis.major_unit {
        write!(writer, r#"<c:majorUnit val="{}"/>"#, major_unit)?;
    }
    if let Some(minor_unit) = axis.minor_unit {
        write!(writer, r#"<c:minorUnit val="{}"/>"#, minor_unit)?;
    }
    if let Some(display_units) = &axis.display_units {
        if display_units.built_in_unit.is_some() == display_units.custom_unit.is_some() {
            return Err(invalid_chart_input(
                "chart display units require exactly one built-in or custom unit",
            ));
        }
        write!(writer, "<c:dispUnits>")?;
        if let Some(unit) = display_units.built_in_unit {
            write!(writer, r#"<c:builtInUnit val="{}"/>"#, unit.xml_value())?;
        }
        if let Some(unit) = display_units.custom_unit {
            if !unit.is_finite() || unit <= 0.0 {
                return Err(invalid_chart_input(
                    "chart custom display unit must be finite and positive",
                ));
            }
            write!(writer, r#"<c:custUnit val="{}"/>"#, unit)?;
        }
        if display_units.show_label
            || display_units.label.is_some()
            || display_units.layout.is_some()
            || display_units.label_shape_properties.is_some()
            || display_units.label_text_properties.is_some()
        {
            write!(writer, "<c:dispUnitsLbl>")?;
            if let Some(layout) = display_units.layout.as_ref() {
                write_layout(writer, Some(layout))?;
            }
            if let Some(label) = display_units.label.as_ref() {
                write_title_text(writer, label)?;
            }
            if let Some(shape_properties) = display_units.label_shape_properties.as_ref() {
                write_fragment(writer, shape_properties.as_xml())?;
            }
            if let Some(text_properties) = display_units.label_text_properties.as_ref() {
                write_fragment(writer, text_properties.as_xml())?;
            }
            write!(writer, "</c:dispUnitsLbl>")?;
        }
        if let Some(extension_list) = display_units.extension_list.as_ref() {
            write_fragment(writer, extension_list.as_xml())?;
        }
        write!(writer, "</c:dispUnits>")?;
    }

    write_axis_extension(writer, &axis.common)?;
    write!(writer, "</c:valAx>")?;
    Ok(())
}

pub(super) fn write_date_axis<W: Write>(writer: &mut W, axis: &DateAxis) -> std::io::Result<()> {
    write!(writer, "<c:dateAx>")?;
    write_axis_common(writer, &axis.common, axis.min, axis.max, axis.log_base)?;
    write!(
        writer,
        r#"<c:auto val="{}"/>"#,
        if axis.auto { "1" } else { "0" }
    )?;
    if let Some(offset) = axis.label_offset {
        if offset > 1000 {
            return Err(invalid_chart_input(
                "chart date-axis label offset must be between 0 and 1000",
            ));
        }
        write!(writer, r#"<c:lblOffset val="{}"/>"#, offset)?;
    }
    if let Some(unit) = axis.base_time_unit {
        write!(writer, r#"<c:baseTimeUnit val="{}"/>"#, unit.xml_value())?;
    }
    if let Some(unit) = axis.major_unit {
        write!(writer, r#"<c:majorUnit val="{}"/>"#, unit)?;
    }
    if let Some(unit) = axis.major_time_unit {
        write!(writer, r#"<c:majorTimeUnit val="{}"/>"#, unit.xml_value())?;
    }
    if let Some(unit) = axis.minor_unit {
        write!(writer, r#"<c:minorUnit val="{}"/>"#, unit)?;
    }
    if let Some(unit) = axis.minor_time_unit {
        write!(writer, r#"<c:minorTimeUnit val="{}"/>"#, unit.xml_value())?;
    }
    write_axis_extension(writer, &axis.common)?;
    write!(writer, "</c:dateAx>")?;
    Ok(())
}

pub(super) fn write_series_axis<W: Write>(
    writer: &mut W,
    axis: &SeriesAxis,
) -> std::io::Result<()> {
    write!(writer, "<c:serAx>")?;
    write_axis_common(writer, &axis.common, axis.min, axis.max, axis.log_base)?;
    if let Some(skip) = axis.tick_label_skip {
        write!(writer, r#"<c:tickLblSkip val="{}"/>"#, skip)?;
    }
    if let Some(skip) = axis.tick_mark_skip {
        write!(writer, r#"<c:tickMarkSkip val="{}"/>"#, skip)?;
    }
    write_axis_extension(writer, &axis.common)?;
    write!(writer, "</c:serAx>")?;
    Ok(())
}

pub(super) fn write_axis_extension<W: Write>(
    writer: &mut W,
    common: &AxisCommon,
) -> std::io::Result<()> {
    if let Some(extension_list) = common.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }
    Ok(())
}
