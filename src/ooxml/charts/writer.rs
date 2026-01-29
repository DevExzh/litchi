//! Chart XML writer.
//!
//! This module provides functionality to generate chart XML for OOXML packages.

use crate::common::xml::escape_xml;
use crate::ooxml::charts::axis::{Axis, AxisCommon, CategoryAxis, DateAxis, SeriesAxis, ValueAxis};
use crate::ooxml::charts::chart::{Chart, View3D, WallFloor};
use crate::ooxml::charts::legend::Legend;
use crate::ooxml::charts::models::{NumericData, StringData, TitleText};
use crate::ooxml::charts::plot_area::{
    Area3DTypeGroup, AreaTypeGroup, Bar3DTypeGroup, BarTypeGroup, BubbleTypeGroup,
    DoughnutTypeGroup, Line3DTypeGroup, LineTypeGroup, Pie3DTypeGroup, PieTypeGroup, PlotArea,
    RadarTypeGroup, ScatterTypeGroup, StockTypeGroup, Surface3DTypeGroup, SurfaceTypeGroup,
    TypeGroup,
};
use crate::ooxml::charts::series::Series;
use std::io::Write;

/// Write a chart to XML.
pub fn write_chart<W: Write>(writer: &mut W, chart: &Chart) -> std::io::Result<()> {
    write!(
        writer,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#
    )?;
    write!(
        writer,
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#
    )?;
    write!(
        writer,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#
    )?;
    write!(
        writer,
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#
    )?;

    write!(
        writer,
        r#"<c:date1904 val="{}"/>"#,
        if chart.date_1904 { "1" } else { "0" }
    )?;
    write!(writer, r#"<c:lang val="en-US"/>"#)?;
    write!(
        writer,
        r#"<c:roundedCorners val="{}"/>"#,
        if chart.rounded_corners { "1" } else { "0" }
    )?;

    if let Some(ref style) = chart.style {
        write!(writer, r#"<c:style val="{}"/>"#, style)?;
    }

    write!(writer, "<c:chart>")?;

    if let Some(ref title) = chart.title {
        write_title(writer, title)?;
    }

    write!(
        writer,
        r#"<c:autoTitleDeleted val="{}"/>"#,
        if chart.auto_title_deleted { "1" } else { "0" }
    )?;

    if let Some(ref view) = chart.view_3d {
        write_view_3d(writer, view)?;
    }

    if let Some(ref floor) = chart.floor {
        write!(writer, "<c:floor>")?;
        write_wall_floor(writer, floor)?;
        write!(writer, "</c:floor>")?;
    }

    if let Some(ref back_wall) = chart.back_wall {
        write!(writer, "<c:backWall>")?;
        write_wall_floor(writer, back_wall)?;
        write!(writer, "</c:backWall>")?;
    }

    if let Some(ref side_wall) = chart.side_wall {
        write!(writer, "<c:sideWall>")?;
        write_wall_floor(writer, side_wall)?;
        write!(writer, "</c:sideWall>")?;
    }

    write_plot_area(writer, &chart.plot_area)?;

    if let Some(ref legend) = chart.legend {
        write_legend(writer, legend)?;
    }

    write!(
        writer,
        r#"<c:plotVisOnly val="{}"/>"#,
        if chart.plot_visible_only { "1" } else { "0" }
    )?;
    write!(
        writer,
        r#"<c:dispBlanksAs val="{}"/>"#,
        chart.display_blanks_as.xml_value()
    )?;

    if chart.show_data_labels_over_max {
        write!(writer, r#"<c:showDLblsOverMax val="1"/>"#)?;
    }

    write!(writer, "</c:chart>")?;

    write!(writer, "<c:printSettings>")?;
    write!(writer, "<c:headerFooter/>")?;
    write!(
        writer,
        r#"<c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>"#
    )?;
    write!(writer, "<c:pageSetup/>")?;
    write!(writer, "</c:printSettings>")?;

    write!(writer, "</c:chartSpace>")?;

    Ok(())
}

fn write_title<W: Write>(writer: &mut W, title: &TitleText) -> std::io::Result<()> {
    write!(writer, "<c:title>")?;

    match title {
        TitleText::Literal(rich_text) => {
            write!(writer, "<c:tx><c:rich>")?;
            write!(writer, "<a:bodyPr/><a:lstStyle/>")?;
            write!(writer, "<a:p><a:pPr><a:defRPr/></a:pPr>")?;
            write!(
                writer,
                r#"<a:r><a:rPr lang="en-US"/><a:t>{}</a:t></a:r>"#,
                escape_xml(&rich_text.text)
            )?;
            write!(writer, "</a:p></c:rich></c:tx>")?;
        },
        TitleText::Reference(source_ref) => {
            write!(writer, "<c:tx><c:strRef>")?;
            write!(writer, "<c:f>{}</c:f>", escape_xml(&source_ref.formula))?;
            write!(writer, "</c:strRef></c:tx>")?;
        },
    }

    write!(writer, r#"<c:overlay val="0"/>"#)?;
    write!(writer, "</c:title>")?;

    Ok(())
}

fn write_view_3d<W: Write>(writer: &mut W, view: &View3D) -> std::io::Result<()> {
    write!(writer, "<c:view3D>")?;

    if let Some(rot_x) = view.rot_x {
        write!(writer, r#"<c:rotX val="{}"/>"#, rot_x)?;
    }
    if let Some(rot_y) = view.rot_y {
        write!(writer, r#"<c:rotY val="{}"/>"#, rot_y)?;
    }

    write!(
        writer,
        r#"<c:rAngAx val="{}"/>"#,
        if view.right_angle_axes { "1" } else { "0" }
    )?;

    if let Some(perspective) = view.perspective {
        write!(writer, r#"<c:perspective val="{}"/>"#, perspective)?;
    }
    if let Some(height) = view.height_percent {
        write!(writer, r#"<c:hPercent val="{}"/>"#, height)?;
    }
    if let Some(depth) = view.depth_percent {
        write!(writer, r#"<c:depthPercent val="{}"/>"#, depth)?;
    }

    write!(writer, "</c:view3D>")?;

    Ok(())
}

fn write_wall_floor<W: Write>(writer: &mut W, wall_floor: &WallFloor) -> std::io::Result<()> {
    if let Some(thickness) = wall_floor.thickness {
        write!(writer, r#"<c:thickness val="{}"/>"#, thickness)?;
    }
    Ok(())
}

fn write_plot_area<W: Write>(writer: &mut W, plot_area: &PlotArea) -> std::io::Result<()> {
    write!(writer, "<c:plotArea>")?;
    write!(writer, "<c:layout/>")?;

    for type_group in &plot_area.type_groups {
        write_type_group(writer, type_group)?;
    }

    for axis in &plot_area.axes {
        write_axis(writer, axis)?;
    }

    write!(writer, "</c:plotArea>")?;

    Ok(())
}

fn write_type_group<W: Write>(writer: &mut W, type_group: &TypeGroup) -> std::io::Result<()> {
    match type_group {
        TypeGroup::Area(group) => write_area_chart(writer, group),
        TypeGroup::Area3D(group) => write_area_3d_chart(writer, group),
        TypeGroup::Bar(group) => write_bar_chart(writer, group),
        TypeGroup::Bar3D(group) => write_bar_3d_chart(writer, group),
        TypeGroup::Bubble(group) => write_bubble_chart(writer, group),
        TypeGroup::Doughnut(group) => write_doughnut_chart(writer, group),
        TypeGroup::Line(group) => write_line_chart(writer, group),
        TypeGroup::Line3D(group) => write_line_3d_chart(writer, group),
        TypeGroup::Pie(group) => write_pie_chart(writer, group),
        TypeGroup::Pie3D(group) => write_pie_3d_chart(writer, group),
        TypeGroup::Radar(group) => write_radar_chart(writer, group),
        TypeGroup::Scatter(group) => write_scatter_chart(writer, group),
        TypeGroup::Stock(group) => write_stock_chart(writer, group),
        TypeGroup::Surface(group) => write_surface_chart(writer, group),
        TypeGroup::Surface3D(group) => write_surface_3d_chart(writer, group),
    }
}

fn write_area_chart<W: Write>(writer: &mut W, group: &AreaTypeGroup) -> std::io::Result<()> {
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
        write_series(writer, series, false)?;
    }

    write_data_labels_default(writer)?;
    write!(writer, r#"<c:axId val="1"/><c:axId val="2"/>"#)?;
    write!(writer, "</c:areaChart>")?;

    Ok(())
}

fn write_area_3d_chart<W: Write>(writer: &mut W, group: &Area3DTypeGroup) -> std::io::Result<()> {
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
        write_series(writer, series, false)?;
    }

    write_data_labels_default(writer)?;
    write!(writer, r#"<c:axId val="1"/><c:axId val="2"/>"#)?;
    write!(writer, "</c:area3DChart>")?;

    Ok(())
}

fn write_bar_chart<W: Write>(writer: &mut W, group: &BarTypeGroup) -> std::io::Result<()> {
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
        write_series(writer, series, false)?;
    }

    write_data_labels_default(writer)?;

    if let Some(gap_width) = group.gap_width {
        write!(writer, r#"<c:gapWidth val="{}"/>"#, gap_width)?;
    } else {
        write!(writer, r#"<c:gapWidth val="150"/>"#)?;
    }

    if let Some(overlap) = group.overlap {
        write!(writer, r#"<c:overlap val="{}"/>"#, overlap)?;
    }

    write!(writer, r#"<c:axId val="1"/><c:axId val="2"/>"#)?;
    write!(writer, "</c:barChart>")?;

    Ok(())
}

fn write_bar_3d_chart<W: Write>(writer: &mut W, group: &Bar3DTypeGroup) -> std::io::Result<()> {
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
        write_series(writer, series, false)?;
    }

    write_data_labels_default(writer)?;

    if let Some(gap_width) = group.gap_width {
        write!(writer, r#"<c:gapWidth val="{}"/>"#, gap_width)?;
    }

    if let Some(gap_depth) = group.gap_depth {
        write!(writer, r#"<c:gapDepth val="{}"/>"#, gap_depth)?;
    }

    if let Some(ref shape) = group.shape {
        write!(writer, r#"<c:shape val="{}"/>"#, shape.xml_value())?;
    }

    write!(writer, r#"<c:axId val="1"/><c:axId val="2"/>"#)?;
    write!(writer, "</c:bar3DChart>")?;

    Ok(())
}

fn write_bubble_chart<W: Write>(writer: &mut W, group: &BubbleTypeGroup) -> std::io::Result<()> {
    write!(writer, "<c:bubbleChart>")?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_bubble_series(writer, series)?;
    }

    write_data_labels_default(writer)?;

    // bubbleScale defaults to 100 if not specified
    let scale = group.bubble_scale.unwrap_or(100);
    write!(writer, r#"<c:bubbleScale val="{}"/>"#, scale)?;

    write!(
        writer,
        r#"<c:showNegBubbles val="{}"/>"#,
        if group.show_negative_bubbles {
            "1"
        } else {
            "0"
        }
    )?;

    write!(writer, r#"<c:axId val="1"/><c:axId val="2"/>"#)?;
    write!(writer, "</c:bubbleChart>")?;

    Ok(())
}

fn write_doughnut_chart<W: Write>(
    writer: &mut W,
    group: &DoughnutTypeGroup,
) -> std::io::Result<()> {
    write!(writer, "<c:doughnutChart>")?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, true)?;
    }

    write_data_labels_default(writer)?;
    write!(
        writer,
        r#"<c:firstSliceAng val="{}"/>"#,
        group.first_slice_angle
    )?;
    write!(writer, r#"<c:holeSize val="{}"/>"#, group.hole_size)?;
    write!(writer, "</c:doughnutChart>")?;

    Ok(())
}

fn write_line_chart<W: Write>(writer: &mut W, group: &LineTypeGroup) -> std::io::Result<()> {
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
        write_series(writer, series, false)?;
    }

    write_data_labels_default(writer)?;
    write!(
        writer,
        r#"<c:marker val="{}"/>"#,
        if group.marker { "1" } else { "0" }
    )?;
    write!(writer, r#"<c:axId val="1"/><c:axId val="2"/>"#)?;
    write!(writer, "</c:lineChart>")?;

    Ok(())
}

fn write_line_3d_chart<W: Write>(writer: &mut W, group: &Line3DTypeGroup) -> std::io::Result<()> {
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
        write_series(writer, series, false)?;
    }

    write_data_labels_default(writer)?;
    write!(writer, r#"<c:axId val="1"/><c:axId val="2"/>"#)?;
    write!(writer, "</c:line3DChart>")?;

    Ok(())
}

fn write_pie_chart<W: Write>(writer: &mut W, group: &PieTypeGroup) -> std::io::Result<()> {
    write!(writer, "<c:pieChart>")?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, true)?;
    }

    write_data_labels_default(writer)?;
    write!(
        writer,
        r#"<c:firstSliceAng val="{}"/>"#,
        group.first_slice_angle
    )?;
    write!(writer, "</c:pieChart>")?;

    Ok(())
}

fn write_pie_3d_chart<W: Write>(writer: &mut W, group: &Pie3DTypeGroup) -> std::io::Result<()> {
    write!(writer, "<c:pie3DChart>")?;
    write!(
        writer,
        r#"<c:varyColors val="{}"/>"#,
        if group.common.vary_colors { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, true)?;
    }

    write_data_labels_default(writer)?;
    write!(writer, "</c:pie3DChart>")?;

    Ok(())
}

fn write_radar_chart<W: Write>(writer: &mut W, group: &RadarTypeGroup) -> std::io::Result<()> {
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
        write_series(writer, series, false)?;
    }

    write_data_labels_default(writer)?;
    write!(writer, r#"<c:axId val="1"/><c:axId val="2"/>"#)?;
    write!(writer, "</c:radarChart>")?;

    Ok(())
}

fn write_scatter_chart<W: Write>(writer: &mut W, group: &ScatterTypeGroup) -> std::io::Result<()> {
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

    write_data_labels_default(writer)?;
    write!(writer, r#"<c:axId val="1"/><c:axId val="2"/>"#)?;
    write!(writer, "</c:scatterChart>")?;

    Ok(())
}

fn write_stock_chart<W: Write>(writer: &mut W, group: &StockTypeGroup) -> std::io::Result<()> {
    write!(writer, "<c:stockChart>")?;

    for series in &group.common.series {
        write_series(writer, series, false)?;
    }

    write_data_labels_default(writer)?;
    write!(writer, r#"<c:axId val="1"/><c:axId val="2"/>"#)?;
    write!(writer, "</c:stockChart>")?;

    Ok(())
}

fn write_surface_chart<W: Write>(writer: &mut W, group: &SurfaceTypeGroup) -> std::io::Result<()> {
    write!(writer, "<c:surfaceChart>")?;
    write!(
        writer,
        r#"<c:wireframe val="{}"/>"#,
        if group.wireframe { "1" } else { "0" }
    )?;

    for series in &group.common.series {
        write_series(writer, series, false)?;
    }

    write!(
        writer,
        r#"<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/>"#
    )?;
    write!(writer, "</c:surfaceChart>")?;

    Ok(())
}

fn write_surface_3d_chart<W: Write>(
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
        write_series(writer, series, false)?;
    }

    write!(
        writer,
        r#"<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/>"#
    )?;
    write!(writer, "</c:surface3DChart>")?;

    Ok(())
}

fn write_series<W: Write>(writer: &mut W, series: &Series, is_pie: bool) -> std::io::Result<()> {
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

    if let Some(ref categories) = series.categories {
        write_string_data_ref(writer, "c:cat", categories)?;
    }

    if let Some(ref values) = series.values {
        write_numeric_data_ref(writer, "c:val", values)?;
    }

    if is_pie && let Some(explosion) = series.explosion {
        write!(writer, r#"<c:explosion val="{}"/>"#, explosion)?;
    }

    write!(writer, "</c:ser>")?;

    Ok(())
}

fn write_scatter_series<W: Write>(writer: &mut W, series: &Series) -> std::io::Result<()> {
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

    if let Some(ref x_values) = series.x_values {
        write_numeric_data_ref(writer, "c:xVal", x_values)?;
    }

    if let Some(ref y_values) = series.y_values {
        write_numeric_data_ref(writer, "c:yVal", y_values)?;
    }

    if let Some(ref bubble_sizes) = series.bubble_sizes {
        write_numeric_data_ref(writer, "c:bubbleSize", bubble_sizes)?;
    }

    write!(writer, "</c:ser>")?;

    Ok(())
}

fn write_bubble_series<W: Write>(writer: &mut W, series: &Series) -> std::io::Result<()> {
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

    // Bubble charts do NOT have xVal - only yVal and bubbleSize
    if let Some(ref y_values) = series.y_values {
        write_numeric_data_ref(writer, "c:yVal", y_values)?;
    }

    if let Some(ref bubble_sizes) = series.bubble_sizes {
        write_numeric_data_ref(writer, "c:bubbleSize", bubble_sizes)?;
    }

    // Per OOXML spec, bubble3D appears in each series
    write!(writer, r#"<c:bubble3D val="0"/>"#)?;

    write!(writer, "</c:ser>")?;

    Ok(())
}

fn write_string_data_ref<W: Write>(
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

fn write_numeric_data_ref<W: Write>(
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

fn write_data_labels_default<W: Write>(writer: &mut W) -> std::io::Result<()> {
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

fn write_axis<W: Write>(writer: &mut W, axis: &Axis) -> std::io::Result<()> {
    match axis {
        Axis::Category(ax) => write_category_axis(writer, ax),
        Axis::Value(ax) => write_value_axis(writer, ax),
        Axis::Date(ax) => write_date_axis(writer, ax),
        Axis::Series(ax) => write_series_axis(writer, ax),
    }
}

fn write_axis_common<W: Write>(writer: &mut W, common: &AxisCommon) -> std::io::Result<()> {
    write!(writer, r#"<c:axId val="{}"/>"#, common.axis_id)?;

    write!(writer, "<c:scaling>")?;
    write!(
        writer,
        r#"<c:orientation val="{}"/>"#,
        common.orientation.xml_value()
    )?;
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

    if common.show_major_gridlines {
        write!(writer, "<c:majorGridlines/>")?;
    }

    if common.show_minor_gridlines {
        write!(writer, "<c:minorGridlines/>")?;
    }

    if let Some(ref title) = common.title {
        write_title(writer, title)?;
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

fn write_category_axis<W: Write>(writer: &mut W, axis: &CategoryAxis) -> std::io::Result<()> {
    write!(writer, "<c:catAx>")?;
    write_axis_common(writer, &axis.common)?;
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
    write!(
        writer,
        r#"<c:noMultiLvlLbl val="{}"/>"#,
        if axis.no_multi_level { "1" } else { "0" }
    )?;
    write!(writer, "</c:catAx>")?;
    Ok(())
}

fn write_value_axis<W: Write>(writer: &mut W, axis: &ValueAxis) -> std::io::Result<()> {
    write!(writer, "<c:valAx>")?;
    write_axis_common(writer, &axis.common)?;

    write!(
        writer,
        r#"<c:crossBetween val="{}"/>"#,
        axis.cross_between.xml_value()
    )?;

    if let Some(min) = axis.min {
        write!(writer, r#"<c:scaling><c:min val="{}"/></c:scaling>"#, min)?;
    }
    if let Some(max) = axis.max {
        write!(writer, r#"<c:scaling><c:max val="{}"/></c:scaling>"#, max)?;
    }
    if let Some(major_unit) = axis.major_unit {
        write!(writer, r#"<c:majorUnit val="{}"/>"#, major_unit)?;
    }
    if let Some(minor_unit) = axis.minor_unit {
        write!(writer, r#"<c:minorUnit val="{}"/>"#, minor_unit)?;
    }

    write!(writer, "</c:valAx>")?;
    Ok(())
}

fn write_date_axis<W: Write>(writer: &mut W, axis: &DateAxis) -> std::io::Result<()> {
    write!(writer, "<c:dateAx>")?;
    write_axis_common(writer, &axis.common)?;
    write!(
        writer,
        r#"<c:auto val="{}"/>"#,
        if axis.auto { "1" } else { "0" }
    )?;
    write!(writer, "</c:dateAx>")?;
    Ok(())
}

fn write_series_axis<W: Write>(writer: &mut W, axis: &SeriesAxis) -> std::io::Result<()> {
    write!(writer, "<c:serAx>")?;
    write_axis_common(writer, &axis.common)?;
    write!(writer, "</c:serAx>")?;
    Ok(())
}

fn write_legend<W: Write>(writer: &mut W, legend: &Legend) -> std::io::Result<()> {
    write!(writer, "<c:legend>")?;
    write!(
        writer,
        r#"<c:legendPos val="{}"/>"#,
        legend.position.xml_value()
    )?;
    write!(
        writer,
        r#"<c:overlay val="{}"/>"#,
        if legend.overlay { "1" } else { "0" }
    )?;
    write!(writer, "</c:legend>")?;
    Ok(())
}
