//! Shared chart-line, extension, axis-ID, and collection records.

use super::super::validation::{invalid_chart_input, validate_optional_u32_range};
use super::super::xml::write_fragment;
use super::series::write_data_labels;
use crate::chart::plot_area::{BandFormat, Lines, TypeGroupCommon, UpDownBars};
use std::io::Write;

pub(super) fn write_data_labels_default<W: Write>(writer: &mut W) -> std::io::Result<()> {
    write!(writer, "<c:dLbls>")?;
    write!(writer, r#"<c:showLegendKey val="0"/>"#)?;
    write!(writer, r#"<c:showVal val="0"/>"#)?;
    write!(writer, r#"<c:showCatName val="0"/>"#)?;
    write!(writer, r#"<c:showSerName val="0"/>"#)?;
    write!(writer, r#"<c:showPercent val="0"/>"#)?;
    write!(writer, r#"<c:showBubbleSize val="0"/>"#)?;
    write!(writer, "</c:dLbls>")?;
    Ok(())
}

pub(super) fn write_group_data_labels<W: Write>(
    writer: &mut W,
    common: &TypeGroupCommon,
) -> std::io::Result<()> {
    if let Some(labels) = common.data_labels.as_ref() {
        write_data_labels(writer, labels)
    } else {
        write_data_labels_default(writer)
    }
}

pub(super) fn write_up_down_bars<W: Write>(
    writer: &mut W,
    bars: &UpDownBars,
) -> std::io::Result<()> {
    validate_optional_u32_range(bars.gap_width, 0, 500, "chart up/down-bar gap width")?;
    write!(writer, "<c:upDownBars>")?;
    if let Some(gap_width) = bars.gap_width {
        write!(writer, r#"<c:gapWidth val="{gap_width}"/>"#)?;
    }
    if let Some(lines) = bars.up_bars.as_ref() {
        write_chart_lines(writer, "upBars", lines)?;
    }
    if let Some(lines) = bars.down_bars.as_ref() {
        write_chart_lines(writer, "downBars", lines)?;
    }
    if let Some(extension_list) = bars.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }
    write!(writer, "</c:upDownBars>")?;
    Ok(())
}

pub(super) fn write_chart_lines<W: Write>(
    writer: &mut W,
    element_name: &str,
    lines: &Lines,
) -> std::io::Result<()> {
    if let Some(shape_properties) = lines.shape_properties.as_ref() {
        write!(writer, "<c:{element_name}>")?;
        write_fragment(writer, shape_properties.as_xml())?;
        write!(writer, "</c:{element_name}>")?;
    } else {
        write!(writer, "<c:{element_name}/>")?;
    }
    Ok(())
}

pub(super) fn write_type_group_extension<W: Write>(
    writer: &mut W,
    common: &TypeGroupCommon,
) -> std::io::Result<()> {
    if let Some(extension_list) = common.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }
    Ok(())
}

pub(super) fn write_surface_band_formats<W: Write>(
    writer: &mut W,
    formats: &[BandFormat],
) -> std::io::Result<()> {
    let mut indexes = std::collections::HashSet::with_capacity(formats.len());
    write!(writer, "<c:bandFmts>")?;
    for format in formats {
        if !indexes.insert(format.index) {
            return Err(invalid_chart_input(format!(
                "surface chart contains duplicate band index {}",
                format.index
            )));
        }
        write!(writer, r#"<c:bandFmt><c:idx val="{}"/>"#, format.index)?;
        if let Some(shape_properties) = format.shape_properties.as_ref() {
            write_fragment(writer, shape_properties.as_xml())?;
        }
        write!(writer, "</c:bandFmt>")?;
    }
    write!(writer, "</c:bandFmts>")?;
    Ok(())
}

pub(super) fn write_type_group_axis_ids<W: Write>(
    writer: &mut W,
    common: &TypeGroupCommon,
    default_ids: &[u32],
    minimum_count: usize,
    maximum_count: usize,
    description: &str,
) -> std::io::Result<()> {
    let axis_ids = if common.axis_ids.is_empty() {
        default_ids
    } else {
        common.axis_ids.as_slice()
    };
    if !(minimum_count..=maximum_count).contains(&axis_ids.len()) {
        return Err(invalid_chart_input(format!(
            "{description} must reference between {minimum_count} and {maximum_count} axes"
        )));
    }
    let mut unique_ids = std::collections::HashSet::with_capacity(axis_ids.len());
    for axis_id in axis_ids {
        if !unique_ids.insert(*axis_id) {
            return Err(invalid_chart_input(format!(
                "{description} contains duplicate axis ID {axis_id}"
            )));
        }
        write!(writer, r#"<c:axId val="{axis_id}"/>"#)?;
    }
    Ok(())
}
