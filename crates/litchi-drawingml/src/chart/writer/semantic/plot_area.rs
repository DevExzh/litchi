//! Plot-area and chart-type record families.

use super::super::model::SeriesFeatures;
use super::super::validation::{invalid_chart_input, validate_optional_u32_range};
use super::super::xml::write_fragment;
use super::{
    axis::write_axis,
    common::{
        write_chart_lines, write_group_data_labels, write_surface_band_formats,
        write_type_group_axis_ids, write_type_group_extension, write_up_down_bars,
    },
    presentation::write_layout,
    series::{write_bubble_series, write_scatter_series, write_series},
};
use crate::chart::plot_area::{
    Area3DTypeGroup, AreaTypeGroup, Bar3DTypeGroup, BarTypeGroup, BubbleTypeGroup, DataTable,
    DoughnutTypeGroup, Line3DTypeGroup, LineTypeGroup, OfPieTypeGroup, Pie3DTypeGroup,
    PieTypeGroup, PlotArea, RadarTypeGroup, ScatterTypeGroup, StockTypeGroup, Surface3DTypeGroup,
    SurfaceTypeGroup, TypeGroup,
};
use std::io::Write;

pub(super) fn write_plot_area<W: Write>(
    writer: &mut W,
    plot_area: &PlotArea,
) -> std::io::Result<()> {
    write!(writer, "<c:plotArea>")?;
    write_layout(writer, plot_area.layout.as_ref())?;

    for type_group in &plot_area.type_groups {
        write_type_group(writer, type_group)?;
    }

    for axis in &plot_area.axes {
        write_axis(writer, axis)?;
    }
    if let Some(data_table) = plot_area.data_table.as_ref() {
        write_data_table(writer, data_table)?;
    }
    if let Some(shape_properties) = plot_area.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    if let Some(extension_list) = plot_area.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }

    write!(writer, "</c:plotArea>")?;

    Ok(())
}

pub(super) fn write_data_table<W: Write>(
    writer: &mut W,
    data_table: &DataTable,
) -> std::io::Result<()> {
    write!(writer, "<c:dTable>")?;
    for (name, show) in [
        ("showHorzBorder", data_table.show_horizontal_border),
        ("showVertBorder", data_table.show_vertical_border),
        ("showOutline", data_table.show_outline),
        ("showKeys", data_table.show_legend_keys),
    ] {
        if show {
            write!(writer, r#"<c:{name} val="1"/>"#)?;
        }
    }
    if let Some(shape_properties) = data_table.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    if let Some(text_properties) = data_table.text_properties.as_ref() {
        write_fragment(writer, text_properties.as_xml())?;
    }
    if let Some(extension_list) = data_table.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }
    write!(writer, "</c:dTable>")?;
    Ok(())
}

pub(super) fn write_type_group<W: Write>(
    writer: &mut W,
    type_group: &TypeGroup,
) -> std::io::Result<()> {
    match type_group {
        TypeGroup::Area(group) => write_area_chart(writer, group),
        TypeGroup::Area3D(group) => write_area_3d_chart(writer, group),
        TypeGroup::Bar(group) => write_bar_chart(writer, group),
        TypeGroup::Bar3D(group) => write_bar_3d_chart(writer, group),
        TypeGroup::Bubble(group) => write_bubble_chart(writer, group),
        TypeGroup::Doughnut(group) => write_doughnut_chart(writer, group),
        TypeGroup::Line(group) => write_line_chart(writer, group),
        TypeGroup::Line3D(group) => write_line_3d_chart(writer, group),
        TypeGroup::OfPie(group) => write_of_pie_chart(writer, group),
        TypeGroup::Pie(group) => write_pie_chart(writer, group),
        TypeGroup::Pie3D(group) => write_pie_3d_chart(writer, group),
        TypeGroup::Radar(group) => write_radar_chart(writer, group),
        TypeGroup::Scatter(group) => write_scatter_chart(writer, group),
        TypeGroup::Stock(group) => write_stock_chart(writer, group),
        TypeGroup::Surface(group) => write_surface_chart(writer, group),
        TypeGroup::Surface3D(group) => write_surface_3d_chart(writer, group),
    }
}

pub(super) fn write_area_chart<W: Write>(
    writer: &mut W,
    group: &AreaTypeGroup,
) -> std::io::Result<()> {
    write!(writer, "<c:areaChart>")?;
    write!(
        writer,
        r#"<c:grouping val="{}"/>"#,
        group.grouping.xml_value()
    )?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::AREA)?;
    }

    write_group_data_labels(writer, &group.common)?;
    if let Some(lines) = group.drop_lines.as_ref() {
        write_chart_lines(writer, "dropLines", lines)?;
    }
    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "area chart")?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:areaChart>")?;

    Ok(())
}

pub(super) fn write_area_3d_chart<W: Write>(
    writer: &mut W,
    group: &Area3DTypeGroup,
) -> std::io::Result<()> {
    validate_optional_u32_range(group.gap_depth, 0, 500, "area 3D chart gap depth")?;
    write!(writer, "<c:area3DChart>")?;
    write!(
        writer,
        r#"<c:grouping val="{}"/>"#,
        group.grouping.xml_value()
    )?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::AREA)?;
    }

    write_group_data_labels(writer, &group.common)?;
    if let Some(lines) = group.drop_lines.as_ref() {
        write_chart_lines(writer, "dropLines", lines)?;
    }
    if let Some(gap_depth) = group.gap_depth {
        write!(writer, r#"<c:gapDepth val="{gap_depth}"/>"#)?;
    }
    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 3, "3D area chart")?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:area3DChart>")?;

    Ok(())
}

pub(super) fn write_bar_chart<W: Write>(
    writer: &mut W,
    group: &BarTypeGroup,
) -> std::io::Result<()> {
    validate_optional_u32_range(group.gap_width, 0, 500, "bar chart gap width")?;
    if group
        .overlap
        .is_some_and(|value| !(-100..=100).contains(&value))
    {
        return Err(invalid_chart_input(
            "bar chart overlap must be between -100 and 100",
        ));
    }
    write!(writer, "<c:barChart>")?;
    write!(
        writer,
        r#"<c:barDir val="{}"/>"#,
        group.direction.xml_value()
    )?;
    write!(
        writer,
        r#"<c:grouping val="{}"/>"#,
        group.grouping.xml_value()
    )?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::BAR)?;
    }

    write_group_data_labels(writer, &group.common)?;

    if let Some(gap_width) = group.gap_width {
        write!(writer, r#"<c:gapWidth val="{}"/>"#, gap_width)?;
    } else {
        write!(writer, r#"<c:gapWidth val="150"/>"#)?;
    }

    if let Some(overlap) = group.overlap {
        write!(writer, r#"<c:overlap val="{}"/>"#, overlap)?;
    }
    for lines in &group.series_lines {
        write_chart_lines(writer, "serLines", lines)?;
    }

    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "bar chart")?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:barChart>")?;

    Ok(())
}

pub(super) fn write_bar_3d_chart<W: Write>(
    writer: &mut W,
    group: &Bar3DTypeGroup,
) -> std::io::Result<()> {
    validate_optional_u32_range(group.gap_width, 0, 500, "bar 3D chart gap width")?;
    validate_optional_u32_range(group.gap_depth, 0, 500, "bar 3D chart gap depth")?;
    write!(writer, "<c:bar3DChart>")?;
    write!(
        writer,
        r#"<c:barDir val="{}"/>"#,
        group.direction.xml_value()
    )?;
    write!(
        writer,
        r#"<c:grouping val="{}"/>"#,
        group.grouping.xml_value()
    )?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::BAR)?;
    }

    write_group_data_labels(writer, &group.common)?;

    if let Some(gap_width) = group.gap_width {
        write!(writer, r#"<c:gapWidth val="{}"/>"#, gap_width)?;
    }

    if let Some(gap_depth) = group.gap_depth {
        write!(writer, r#"<c:gapDepth val="{}"/>"#, gap_depth)?;
    }

    if let Some(ref shape) = group.shape {
        write!(writer, r#"<c:shape val="{}"/>"#, shape.xml_value())?;
    }

    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 3, "3D bar chart")?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:bar3DChart>")?;

    Ok(())
}

pub(super) fn write_bubble_chart<W: Write>(
    writer: &mut W,
    group: &BubbleTypeGroup,
) -> std::io::Result<()> {
    write!(writer, "<c:bubbleChart>")?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_bubble_series(writer, series)?;
    }

    write_group_data_labels(writer, &group.common)?;

    write!(writer, r#"<c:bubbleScale val="{}"/>"#, group.scale().get())?;

    write!(
        writer,
        r#"<c:showNegBubbles val="{}"/>"#,
        if group.show_negative_bubbles {
            "1"
        } else {
            "0"
        }
    )?;
    write!(
        writer,
        r#"<c:sizeRepresents val="{}"/>"#,
        group.size().xml_value()
    )?;

    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "bubble chart")?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:bubbleChart>")?;

    Ok(())
}

pub(super) fn write_doughnut_chart<W: Write>(
    writer: &mut W,
    group: &DoughnutTypeGroup,
) -> std::io::Result<()> {
    if group.first_slice_angle > 360 {
        return Err(invalid_chart_input(
            "doughnut chart first-slice angle must be between 0 and 360",
        ));
    }
    if !(1..=90).contains(&group.hole_size) {
        return Err(invalid_chart_input(
            "doughnut chart hole size must be between 1 and 90",
        ));
    }
    write!(writer, "<c:doughnutChart>")?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::PIE)?;
    }

    write_group_data_labels(writer, &group.common)?;
    write!(
        writer,
        r#"<c:firstSliceAng val="{}"/>"#,
        group.first_slice_angle
    )?;
    write!(writer, r#"<c:holeSize val="{}"/>"#, group.hole_size)?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:doughnutChart>")?;

    Ok(())
}

pub(super) fn write_line_chart<W: Write>(
    writer: &mut W,
    group: &LineTypeGroup,
) -> std::io::Result<()> {
    write!(writer, "<c:lineChart>")?;
    write!(
        writer,
        r#"<c:grouping val="{}"/>"#,
        group.grouping.xml_value()
    )?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::LINE)?;
    }

    write_group_data_labels(writer, &group.common)?;
    if let Some(lines) = group.drop_lines.as_ref() {
        write_chart_lines(writer, "dropLines", lines)?;
    }
    if let Some(lines) = group.high_low_lines.as_ref() {
        write_chart_lines(writer, "hiLowLines", lines)?;
    }
    if let Some(up_down_bars) = group.up_down_bars.as_ref() {
        write_up_down_bars(writer, up_down_bars)?;
    }
    write!(
        writer,
        r#"<c:marker val="{}"/>"#,
        if group.marker { "1" } else { "0" }
    )?;
    write!(
        writer,
        r#"<c:smooth val="{}"/>"#,
        if group.smooth { "1" } else { "0" }
    )?;
    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "line chart")?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:lineChart>")?;

    Ok(())
}

pub(super) fn write_line_3d_chart<W: Write>(
    writer: &mut W,
    group: &Line3DTypeGroup,
) -> std::io::Result<()> {
    validate_optional_u32_range(group.gap_depth, 0, 500, "line 3D chart gap depth")?;
    write!(writer, "<c:line3DChart>")?;
    write!(
        writer,
        r#"<c:grouping val="{}"/>"#,
        group.grouping.xml_value()
    )?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::LINE_3D)?;
    }

    write_group_data_labels(writer, &group.common)?;
    if let Some(lines) = group.drop_lines.as_ref() {
        write_chart_lines(writer, "dropLines", lines)?;
    }
    if let Some(gap_depth) = group.gap_depth {
        write!(writer, r#"<c:gapDepth val="{gap_depth}"/>"#)?;
    }
    write_type_group_axis_ids(writer, &group.common, &[1, 2, 3], 3, 3, "3D line chart")?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:line3DChart>")?;

    Ok(())
}

pub(super) fn write_pie_chart<W: Write>(
    writer: &mut W,
    group: &PieTypeGroup,
) -> std::io::Result<()> {
    if group.first_slice_angle > 360 {
        return Err(invalid_chart_input(
            "pie chart first-slice angle must be between 0 and 360",
        ));
    }
    write!(writer, "<c:pieChart>")?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::PIE)?;
    }

    write_group_data_labels(writer, &group.common)?;
    write!(
        writer,
        r#"<c:firstSliceAng val="{}"/>"#,
        group.first_slice_angle
    )?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:pieChart>")?;

    Ok(())
}

pub(super) fn write_of_pie_chart<W: Write>(
    writer: &mut W,
    group: &OfPieTypeGroup,
) -> std::io::Result<()> {
    if group.gap_width.is_some_and(|value| value > 500) {
        return Err(invalid_chart_input(
            "of-pie chart gap width must be between 0 and 500",
        ));
    }
    if group
        .second_pie_size
        .is_some_and(|value| !(5..=200).contains(&value))
    {
        return Err(invalid_chart_input(
            "of-pie chart secondary size must be between 5 and 200",
        ));
    }
    if group.split_position.is_some_and(|value| !value.is_finite()) {
        return Err(invalid_chart_input(
            "of-pie chart split position must be finite",
        ));
    }

    write!(writer, "<c:ofPieChart>")?;
    write!(
        writer,
        r#"<c:ofPieType val="{}"/>"#,
        group.of_pie_type.xml_value()
    )?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;
    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::PIE)?;
    }
    write_group_data_labels(writer, &group.common)?;
    if let Some(gap_width) = group.gap_width {
        write!(writer, r#"<c:gapWidth val="{gap_width}"/>"#)?;
    }
    if let Some(split_type) = group.split_type {
        write!(writer, r#"<c:splitType val="{}"/>"#, split_type.xml_value())?;
    }
    if let Some(split_position) = group.split_position {
        write!(writer, r#"<c:splitPos val="{split_position}"/>"#)?;
    }
    if let Some(points) = group.custom_split_points.as_ref() {
        write!(writer, "<c:custSplit>")?;
        for point in points {
            write!(writer, r#"<c:secondPiePt val="{point}"/>"#)?;
        }
        write!(writer, "</c:custSplit>")?;
    }
    if let Some(second_pie_size) = group.second_pie_size {
        write!(writer, r#"<c:secondPieSize val="{second_pie_size}"/>"#)?;
    }
    for lines in &group.series_lines {
        write_chart_lines(writer, "serLines", lines)?;
    }
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:ofPieChart>")?;
    Ok(())
}

pub(super) fn write_pie_3d_chart<W: Write>(
    writer: &mut W,
    group: &Pie3DTypeGroup,
) -> std::io::Result<()> {
    write!(writer, "<c:pie3DChart>")?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::PIE)?;
    }

    write_group_data_labels(writer, &group.common)?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:pie3DChart>")?;

    Ok(())
}

pub(super) fn write_radar_chart<W: Write>(
    writer: &mut W,
    group: &RadarTypeGroup,
) -> std::io::Result<()> {
    write!(writer, "<c:radarChart>")?;
    write!(
        writer,
        r#"<c:radarStyle val="{}"/>"#,
        group.style.xml_value()
    )?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::RADAR)?;
    }

    write_group_data_labels(writer, &group.common)?;
    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "radar chart")?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:radarChart>")?;

    Ok(())
}

pub(super) fn write_scatter_chart<W: Write>(
    writer: &mut W,
    group: &ScatterTypeGroup,
) -> std::io::Result<()> {
    write!(writer, "<c:scatterChart>")?;
    write!(
        writer,
        r#"<c:scatterStyle val="{}"/>"#,
        group.style.xml_value()
    )?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_scatter_series(writer, series)?;
    }

    write_group_data_labels(writer, &group.common)?;
    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "scatter chart")?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:scatterChart>")?;

    Ok(())
}

pub(super) fn write_stock_chart<W: Write>(
    writer: &mut W,
    group: &StockTypeGroup,
) -> std::io::Result<()> {
    write!(writer, "<c:stockChart>")?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::LINE)?;
    }

    write_group_data_labels(writer, &group.common)?;
    if let Some(lines) = group.drop_lines.as_ref() {
        write_chart_lines(writer, "dropLines", lines)?;
    }
    if let Some(lines) = group.high_low_lines.as_ref() {
        write_chart_lines(writer, "hiLowLines", lines)?;
    }
    if let Some(up_down_bars) = group.up_down_bars.as_ref() {
        write_up_down_bars(writer, up_down_bars)?;
    }
    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "stock chart")?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:stockChart>")?;

    Ok(())
}

pub(super) fn write_surface_chart<W: Write>(
    writer: &mut W,
    group: &SurfaceTypeGroup,
) -> std::io::Result<()> {
    write!(writer, "<c:surfaceChart>")?;
    write!(
        writer,
        r#"<c:wireframe val="{}"/>"#,
        if group.wireframe { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::SURFACE)?;
    }
    if let Some(formats) = group.band_formats.as_deref() {
        write_surface_band_formats(writer, formats)?;
    }

    write_type_group_axis_ids(writer, &group.common, &[1, 2, 3], 3, 3, "surface chart")?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:surfaceChart>")?;

    Ok(())
}

pub(super) fn write_surface_3d_chart<W: Write>(
    writer: &mut W,
    group: &Surface3DTypeGroup,
) -> std::io::Result<()> {
    write!(writer, "<c:surface3DChart>")?;
    write!(
        writer,
        r#"<c:wireframe val="{}"/>"#,
        if group.wireframe { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::SURFACE)?;
    }
    if let Some(formats) = group.band_formats.as_deref() {
        write_surface_band_formats(writer, formats)?;
    }

    write_type_group_axis_ids(writer, &group.common, &[1, 2, 3], 3, 3, "3D surface chart")?;
    write_type_group_extension(writer, &group.common)?;
    write!(writer, "</c:surface3DChart>")?;

    Ok(())
}
