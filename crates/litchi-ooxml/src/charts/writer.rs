//! Chart XML writer.
//!
//! This module provides functionality to generate chart XML for OOXML packages.

use crate::charts::axis::{Axis, AxisCommon, CategoryAxis, DateAxis, SeriesAxis, ValueAxis};
use crate::charts::chart::{
    Chart, ChartHeaderFooter, ChartPageMargins, ChartPageSetup, ChartPrintSettings, PivotFormat,
    View3D, WallFloor,
};
use crate::charts::legend::Legend;
use crate::charts::models::{Layout, NumericData, StringData, TitleText};
use crate::charts::plot_area::{
    Area3DTypeGroup, AreaTypeGroup, BandFormat, Bar3DTypeGroup, BarTypeGroup, BubbleTypeGroup,
    DataTable, DoughnutTypeGroup, Line3DTypeGroup, LineTypeGroup, OfPieTypeGroup, Pie3DTypeGroup,
    PieTypeGroup, PlotArea, RadarTypeGroup, ScatterTypeGroup, StockTypeGroup, Surface3DTypeGroup,
    SurfaceTypeGroup, TypeGroup, TypeGroupCommon, UpDownBars,
};
use crate::charts::series::{
    DataLabel, DataLabels, DataPoint, ErrorBar, ErrorBarDirection, ErrorBarType, ErrorBarValueType,
    Marker, Series, Trendline, TrendlineType,
};
use litchi_core::xml::escape_xml;
use std::io::Write;

#[derive(Clone, Copy)]
struct SeriesFeatures {
    point_and_label_overrides: bool,
    error_bars: bool,
    trendlines: bool,
    explosion: bool,
    invert_if_negative: bool,
    marker: bool,
    smooth: bool,
}

impl SeriesFeatures {
    const BASIC: Self = Self {
        point_and_label_overrides: true,
        error_bars: true,
        trendlines: true,
        explosion: false,
        invert_if_negative: false,
        marker: false,
        smooth: false,
    };
    const AREA: Self = Self {
        invert_if_negative: true,
        ..Self::BASIC
    };
    const BAR: Self = Self {
        invert_if_negative: true,
        ..Self::BASIC
    };
    const LINE: Self = Self {
        marker: true,
        smooth: true,
        ..Self::BASIC
    };
    const LINE_3D: Self = Self {
        marker: true,
        ..Self::BASIC
    };
    const PIE: Self = Self {
        explosion: true,
        error_bars: false,
        trendlines: false,
        ..Self::BASIC
    };
    const RADAR: Self = Self {
        marker: true,
        error_bars: false,
        trendlines: false,
        ..Self::BASIC
    };
    const SURFACE: Self = Self {
        point_and_label_overrides: false,
        error_bars: false,
        trendlines: false,
        ..Self::BASIC
    };
}

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
        write_title(
            writer,
            title,
            chart.title_layout.as_ref(),
            chart.title_overlay,
        )?;
    }

    write!(
        writer,
        r#"<c:autoTitleDeleted val="{}"/>"#,
        if chart.auto_title_deleted { "1" } else { "0" }
    )?;

    if let Some(formats) = chart.pivot_formats.as_deref() {
        write_pivot_formats(writer, formats)?;
    }

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

    if let Some(settings) = chart.print_settings.as_ref() {
        write_print_settings(writer, settings)?;
    }

    write!(writer, "</c:chartSpace>")?;

    Ok(())
}

fn write_title<W: Write>(
    writer: &mut W,
    title: &TitleText,
    layout: Option<&Layout>,
    overlay: bool,
) -> std::io::Result<()> {
    write!(writer, "<c:title>")?;

    write_title_text(writer, title)?;

    if let Some(layout) = layout {
        write_layout(writer, Some(layout))?;
    }

    write!(
        writer,
        r#"<c:overlay val="{}"/>"#,
        if overlay { "1" } else { "0" }
    )?;
    write!(writer, "</c:title>")?;

    Ok(())
}

fn write_pivot_formats<W: Write>(writer: &mut W, formats: &[PivotFormat]) -> std::io::Result<()> {
    let mut indexes = std::collections::HashSet::with_capacity(formats.len());
    write!(writer, "<c:pivotFmts>")?;
    for format in formats {
        if !indexes.insert(format.index) {
            return Err(invalid_chart_input(format!(
                "chart contains duplicate pivot-format index {}",
                format.index
            )));
        }
        write!(writer, r#"<c:pivotFmt><c:idx val="{}"/>"#, format.index)?;
        if let Some(marker) = format.marker.as_ref() {
            write_marker(writer, marker, "chart pivot-format")?;
        }
        if let Some(label) = format.data_label.as_ref() {
            write_data_label(writer, label)?;
        }
        write!(writer, "</c:pivotFmt>")?;
    }
    write!(writer, "</c:pivotFmts>")?;
    Ok(())
}

fn write_print_settings<W: Write>(
    writer: &mut W,
    settings: &ChartPrintSettings,
) -> std::io::Result<()> {
    write!(writer, "<c:printSettings>")?;
    if let Some(header_footer) = settings.header_footer.as_ref() {
        write_chart_header_footer(writer, header_footer)?;
    }
    if let Some(margins) = settings.page_margins.as_ref() {
        write_chart_page_margins(writer, margins)?;
    }
    if let Some(setup) = settings.page_setup.as_ref() {
        write_chart_page_setup(writer, setup)?;
    }
    write!(writer, "</c:printSettings>")?;
    Ok(())
}

fn write_chart_header_footer<W: Write>(
    writer: &mut W,
    header_footer: &ChartHeaderFooter,
) -> std::io::Result<()> {
    write!(
        writer,
        r#"<c:headerFooter alignWithMargins="{}" differentOddEven="{}" differentFirst="{}">"#,
        if header_footer.align_with_margins {
            "1"
        } else {
            "0"
        },
        if header_footer.different_odd_even {
            "1"
        } else {
            "0"
        },
        if header_footer.different_first {
            "1"
        } else {
            "0"
        }
    )?;
    for (name, value) in [
        ("oddHeader", header_footer.odd_header.as_ref()),
        ("oddFooter", header_footer.odd_footer.as_ref()),
        ("evenHeader", header_footer.even_header.as_ref()),
        ("evenFooter", header_footer.even_footer.as_ref()),
        ("firstHeader", header_footer.first_header.as_ref()),
        ("firstFooter", header_footer.first_footer.as_ref()),
    ] {
        if let Some(value) = value {
            write!(writer, "<c:{name}>{}</c:{name}>", escape_xml(value))?;
        }
    }
    write!(writer, "</c:headerFooter>")?;
    Ok(())
}

fn write_chart_page_margins<W: Write>(
    writer: &mut W,
    margins: &ChartPageMargins,
) -> std::io::Result<()> {
    for (name, value) in [
        ("left", margins.left),
        ("right", margins.right),
        ("top", margins.top),
        ("bottom", margins.bottom),
        ("header", margins.header),
        ("footer", margins.footer),
    ] {
        if !value.is_finite() {
            return Err(invalid_chart_input(format!(
                "chart {name} page margin must be finite"
            )));
        }
    }
    write!(
        writer,
        r#"<c:pageMargins l="{}" r="{}" t="{}" b="{}" header="{}" footer="{}"/>"#,
        margins.left, margins.right, margins.top, margins.bottom, margins.header, margins.footer
    )?;
    Ok(())
}

fn write_chart_page_setup<W: Write>(writer: &mut W, setup: &ChartPageSetup) -> std::io::Result<()> {
    write!(
        writer,
        r#"<c:pageSetup paperSize="{}" firstPageNumber="{}" orientation="{}" blackAndWhite="{}" draft="{}" useFirstPageNumber="{}" horizontalDpi="{}" verticalDpi="{}" copies="{}"/>"#,
        setup.paper_size,
        setup.first_page_number,
        setup.orientation.xml_value(),
        if setup.black_and_white { "1" } else { "0" },
        if setup.draft { "1" } else { "0" },
        if setup.use_first_page_number {
            "1"
        } else {
            "0"
        },
        setup.horizontal_dpi,
        setup.vertical_dpi,
        setup.copies
    )?;
    Ok(())
}

fn write_marker<W: Write>(
    writer: &mut W,
    marker: &Marker,
    description: &str,
) -> std::io::Result<()> {
    if marker.size.is_some_and(|size| !(2..=72).contains(&size)) {
        return Err(invalid_chart_input(format!(
            "{description} marker size must be 2-72"
        )));
    }
    write!(writer, "<c:marker>")?;
    if let Some(symbol) = marker.symbol {
        write!(writer, r#"<c:symbol val="{}"/>"#, symbol.xml_value())?;
    }
    if let Some(size) = marker.size {
        write!(writer, r#"<c:size val="{size}"/>"#)?;
    }
    write!(writer, "</c:marker>")?;
    Ok(())
}

fn write_title_text<W: Write>(writer: &mut W, title: &TitleText) -> std::io::Result<()> {
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

    write!(writer, "</c:plotArea>")?;

    Ok(())
}

fn write_data_table<W: Write>(writer: &mut W, data_table: &DataTable) -> std::io::Result<()> {
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
    write!(writer, "</c:dTable>")?;
    Ok(())
}

fn write_layout<W: Write>(writer: &mut W, layout: Option<&Layout>) -> std::io::Result<()> {
    let Some(layout) = layout else {
        return write!(writer, "<c:layout/>");
    };
    for (name, value) in [
        ("x", layout.x),
        ("y", layout.y),
        ("width", layout.width),
        ("height", layout.height),
    ] {
        if value.is_some_and(|value| !value.is_finite()) {
            return Err(invalid_chart_input(format!(
                "chart layout {name} must be finite"
            )));
        }
    }
    write!(writer, "<c:layout><c:manualLayout>")?;
    if let Some(target) = layout.target {
        write!(writer, r#"<c:layoutTarget val="{}"/>"#, target.xml_value())?;
    }
    for (name, mode) in [
        ("xMode", layout.x_mode),
        ("yMode", layout.y_mode),
        ("wMode", layout.width_mode),
        ("hMode", layout.height_mode),
    ] {
        if let Some(mode) = mode {
            write!(writer, r#"<c:{name} val="{}"/>"#, mode.xml_value())?;
        }
    }
    for (name, value) in [
        ("x", layout.x),
        ("y", layout.y),
        ("w", layout.width),
        ("h", layout.height),
    ] {
        if let Some(value) = value {
            write!(writer, r#"<c:{name} val="{value}"/>"#)?;
        }
    }
    write!(writer, "</c:manualLayout></c:layout>")?;
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
        write_series(writer, series, SeriesFeatures::AREA)?;
    }

    write_group_data_labels(writer, &group.common)?;
    if group.drop_lines.is_some() {
        write!(writer, "<c:dropLines/>")?;
    }
    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "area chart")?;
    write!(writer, "</c:areaChart>")?;

    Ok(())
}

fn write_area_3d_chart<W: Write>(writer: &mut W, group: &Area3DTypeGroup) -> std::io::Result<()> {
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
        write_series(writer, series, SeriesFeatures::BASIC)?;
    }

    write_group_data_labels(writer, &group.common)?;
    if group.drop_lines.is_some() {
        write!(writer, "<c:dropLines/>")?;
    }
    if let Some(gap_depth) = group.gap_depth {
        write!(writer, r#"<c:gapDepth val="{gap_depth}"/>"#)?;
    }
    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 3, "3D area chart")?;
    write!(writer, "</c:area3DChart>")?;

    Ok(())
}

fn write_bar_chart<W: Write>(writer: &mut W, group: &BarTypeGroup) -> std::io::Result<()> {
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
    for _ in &group.series_lines {
        write!(writer, "<c:serLines/>")?;
    }

    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "bar chart")?;
    write!(writer, "</c:barChart>")?;

    Ok(())
}

fn write_bar_3d_chart<W: Write>(writer: &mut W, group: &Bar3DTypeGroup) -> std::io::Result<()> {
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
    write!(writer, "</c:bar3DChart>")?;

    Ok(())
}

fn write_bubble_chart<W: Write>(writer: &mut W, group: &BubbleTypeGroup) -> std::io::Result<()> {
    validate_optional_u32_range(group.bubble_scale, 0, 300, "bubble chart scale")?;
    if !matches!(group.size_represents.as_str(), "area" | "w") {
        return Err(invalid_chart_input(
            "bubble chart size representation must be 'area' or 'w'",
        ));
    }
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

    write!(
        writer,
        r#"<c:bubble3D val="{}"/>"#,
        if group.bubble_3d { "1" } else { "0" }
    )?;

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
    write!(
        writer,
        r#"<c:sizeRepresents val="{}"/>"#,
        escape_xml(&group.size_represents)
    )?;

    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "bubble chart")?;
    write!(writer, "</c:bubbleChart>")?;

    Ok(())
}

fn write_doughnut_chart<W: Write>(
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
        write_series(writer, series, SeriesFeatures::LINE)?;
    }

    write_group_data_labels(writer, &group.common)?;
    if group.drop_lines.is_some() {
        write!(writer, "<c:dropLines/>")?;
    }
    if group.high_low_lines.is_some() {
        write!(writer, "<c:hiLowLines/>")?;
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
    write!(writer, "</c:lineChart>")?;

    Ok(())
}

fn write_line_3d_chart<W: Write>(writer: &mut W, group: &Line3DTypeGroup) -> std::io::Result<()> {
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
    if group.drop_lines.is_some() {
        write!(writer, "<c:dropLines/>")?;
    }
    if let Some(gap_depth) = group.gap_depth {
        write!(writer, r#"<c:gapDepth val="{gap_depth}"/>"#)?;
    }
    write_type_group_axis_ids(writer, &group.common, &[1, 2, 3], 3, 3, "3D line chart")?;
    write!(writer, "</c:line3DChart>")?;

    Ok(())
}

fn write_pie_chart<W: Write>(writer: &mut W, group: &PieTypeGroup) -> std::io::Result<()> {
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
    write!(writer, "</c:pieChart>")?;

    Ok(())
}

fn write_of_pie_chart<W: Write>(writer: &mut W, group: &OfPieTypeGroup) -> std::io::Result<()> {
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
    for _ in &group.series_lines {
        write!(writer, "<c:serLines/>")?;
    }
    write!(writer, "</c:ofPieChart>")?;
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
        write_series(writer, series, SeriesFeatures::PIE)?;
    }

    write_group_data_labels(writer, &group.common)?;
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
        write_series(writer, series, SeriesFeatures::RADAR)?;
    }

    write_group_data_labels(writer, &group.common)?;
    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "radar chart")?;
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

    write_group_data_labels(writer, &group.common)?;
    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "scatter chart")?;
    write!(writer, "</c:scatterChart>")?;

    Ok(())
}

fn write_stock_chart<W: Write>(writer: &mut W, group: &StockTypeGroup) -> std::io::Result<()> {
    write!(writer, "<c:stockChart>")?;

    for series in &group.common.series {
        write_series(writer, series, SeriesFeatures::LINE)?;
    }

    write_group_data_labels(writer, &group.common)?;
    if group.drop_lines.is_some() {
        write!(writer, "<c:dropLines/>")?;
    }
    if group.high_low_lines.is_some() {
        write!(writer, "<c:hiLowLines/>")?;
    }
    if let Some(up_down_bars) = group.up_down_bars.as_ref() {
        write_up_down_bars(writer, up_down_bars)?;
    }
    write_type_group_axis_ids(writer, &group.common, &[1, 2], 2, 2, "stock chart")?;
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
        write_series(writer, series, SeriesFeatures::SURFACE)?;
    }
    if let Some(formats) = group.band_formats.as_deref() {
        write_surface_band_formats(writer, formats)?;
    }

    write_type_group_axis_ids(writer, &group.common, &[1, 2, 3], 3, 3, "surface chart")?;
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
        write_series(writer, series, SeriesFeatures::SURFACE)?;
    }
    if let Some(formats) = group.band_formats.as_deref() {
        write_surface_band_formats(writer, formats)?;
    }

    write_type_group_axis_ids(writer, &group.common, &[1, 2, 3], 3, 3, "3D surface chart")?;
    write!(writer, "</c:surface3DChart>")?;

    Ok(())
}

fn write_series<W: Write>(
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

    write_series_presentation(writer, series, features)?;

    if let Some(ref categories) = series.categories {
        write_string_data_ref(writer, "c:cat", categories)?;
    }

    if let Some(ref values) = series.values {
        write_numeric_data_ref(writer, "c:val", values)?;
    }

    if features.smooth {
        write!(
            writer,
            r#"<c:smooth val="{}"/>"#,
            if series.smooth { "1" } else { "0" }
        )?;
    }

    write!(writer, "</c:ser>")?;

    Ok(())
}

fn write_series_presentation<W: Write>(
    writer: &mut W,
    series: &Series,
    features: SeriesFeatures,
) -> std::io::Result<()> {
    if !features.marker && (series.marker_symbol.is_some() || series.marker_size.is_some()) {
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
    if features.marker && (series.marker_symbol.is_some() || series.marker_size.is_some()) {
        write!(writer, "<c:marker>")?;
        if let Some(symbol) = series.marker_symbol {
            write!(writer, r#"<c:symbol val="{}"/>"#, symbol.xml_value())?;
        }
        if let Some(size) = series.marker_size {
            write!(writer, r#"<c:size val="{}"/>"#, size)?;
        }
        write!(writer, "</c:marker>")?;
    }

    if features.invert_if_negative && series.invert_if_negative {
        write!(writer, r#"<c:invertIfNegative val="1"/>"#)?;
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

fn write_data_point<W: Write>(writer: &mut W, point: &DataPoint) -> std::io::Result<()> {
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
    if point.marker_symbol.is_some() || point.marker_size.is_some() {
        write!(writer, "<c:marker>")?;
        if let Some(symbol) = point.marker_symbol {
            write!(writer, r#"<c:symbol val="{}"/>"#, symbol.xml_value())?;
        }
        if let Some(size) = point.marker_size {
            write!(writer, r#"<c:size val="{}"/>"#, size)?;
        }
        write!(writer, "</c:marker>")?;
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
    write!(writer, "</c:dPt>")?;
    Ok(())
}

fn invalid_chart_input(message: impl Into<String>) -> std::io::Error {
    std::io::Error::new(std::io::ErrorKind::InvalidInput, message.into())
}

fn validate_optional_u32_range(
    value: Option<u32>,
    minimum: u32,
    maximum: u32,
    description: &str,
) -> std::io::Result<()> {
    if value.is_some_and(|value| !(minimum..=maximum).contains(&value)) {
        return Err(invalid_chart_input(format!(
            "{description} must be between {minimum} and {maximum}"
        )));
    }
    Ok(())
}

fn write_data_labels<W: Write>(writer: &mut W, labels: &DataLabels) -> std::io::Result<()> {
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
        write!(writer, r#"<c:delete val="1"/>"#)?;
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
    write!(writer, "</c:dLbls>")?;
    Ok(())
}

fn write_data_label<W: Write>(writer: &mut W, label: &DataLabel) -> std::io::Result<()> {
    write!(writer, r#"<c:dLbl><c:idx val="{}"/>"#, label.index)?;
    if label.deleted {
        write!(writer, r#"<c:delete val="1"/></c:dLbl>"#)?;
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
    write!(writer, "</c:dLbl>")?;
    Ok(())
}

fn write_trendline<W: Write>(writer: &mut W, trendline: &Trendline) -> std::io::Result<()> {
    validate_trendline(trendline)?;
    write!(writer, "<c:trendline>")?;
    if let Some(name) = &trendline.name {
        write!(
            writer,
            "<c:trendlineName>{}</c:trendlineName>",
            escape_xml(name)
        )?;
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
        write!(writer, "</c:trendlineLbl>")?;
    }
    write!(writer, "</c:trendline>")?;
    Ok(())
}

fn validate_trendline(trendline: &Trendline) -> std::io::Result<()> {
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

fn write_error_bar<W: Write>(writer: &mut W, error_bar: &ErrorBar) -> std::io::Result<()> {
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
    write!(writer, "</c:errBars>")?;
    Ok(())
}

fn validate_error_bar(error_bar: &ErrorBar) -> std::io::Result<()> {
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

fn write_scatter_series<W: Write>(writer: &mut W, series: &Series) -> std::io::Result<()> {
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

    write_series_presentation(writer, series, SeriesFeatures::BASIC)?;

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

fn write_group_data_labels<W: Write>(
    writer: &mut W,
    common: &TypeGroupCommon,
) -> std::io::Result<()> {
    if let Some(labels) = common.data_labels.as_ref() {
        write_data_labels(writer, labels)
    } else {
        write_data_labels_default(writer)
    }
}

fn write_up_down_bars<W: Write>(writer: &mut W, bars: &UpDownBars) -> std::io::Result<()> {
    validate_optional_u32_range(bars.gap_width, 0, 500, "chart up/down-bar gap width")?;
    write!(writer, "<c:upDownBars>")?;
    if let Some(gap_width) = bars.gap_width {
        write!(writer, r#"<c:gapWidth val="{gap_width}"/>"#)?;
    }
    if bars.up_bars.is_some() {
        write!(writer, "<c:upBars/>")?;
    }
    if bars.down_bars.is_some() {
        write!(writer, "<c:downBars/>")?;
    }
    write!(writer, "</c:upDownBars>")?;
    Ok(())
}

fn write_surface_band_formats<W: Write>(
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
        write!(
            writer,
            r#"<c:bandFmt><c:idx val="{}"/><c:spPr/></c:bandFmt>"#,
            format.index
        )?;
    }
    write!(writer, "</c:bandFmts>")?;
    Ok(())
}

fn write_type_group_axis_ids<W: Write>(
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

fn write_axis<W: Write>(writer: &mut W, axis: &Axis) -> std::io::Result<()> {
    match axis {
        Axis::Category(ax) => write_category_axis(writer, ax),
        Axis::Value(ax) => write_value_axis(writer, ax),
        Axis::Date(ax) => write_date_axis(writer, ax),
        Axis::Series(ax) => write_series_axis(writer, ax),
    }
}

fn write_axis_common<W: Write>(
    writer: &mut W,
    common: &AxisCommon,
    min: Option<f64>,
    max: Option<f64>,
    log_base: Option<f64>,
) -> std::io::Result<()> {
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
        write_title(writer, title, common.layout.as_ref(), common.title_overlay)?;
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
    write_axis_common(writer, &axis.common, None, None, None)?;
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
    write!(writer, "</c:catAx>")?;
    Ok(())
}

fn write_value_axis<W: Write>(writer: &mut W, axis: &ValueAxis) -> std::io::Result<()> {
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
        {
            write!(writer, "<c:dispUnitsLbl>")?;
            if let Some(layout) = display_units.layout.as_ref() {
                write_layout(writer, Some(layout))?;
            }
            if let Some(label) = display_units.label.as_ref() {
                write_title_text(writer, label)?;
            }
            write!(writer, "</c:dispUnitsLbl>")?;
        }
        write!(writer, "</c:dispUnits>")?;
    }

    write!(writer, "</c:valAx>")?;
    Ok(())
}

fn write_date_axis<W: Write>(writer: &mut W, axis: &DateAxis) -> std::io::Result<()> {
    write!(writer, "<c:dateAx>")?;
    write_axis_common(writer, &axis.common, axis.min, axis.max, None)?;
    write!(
        writer,
        r#"<c:auto val="{}"/>"#,
        if axis.auto { "1" } else { "0" }
    )?;
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
    write!(writer, "</c:dateAx>")?;
    Ok(())
}

fn write_series_axis<W: Write>(writer: &mut W, axis: &SeriesAxis) -> std::io::Result<()> {
    write!(writer, "<c:serAx>")?;
    write_axis_common(writer, &axis.common, None, None, None)?;
    if let Some(skip) = axis.tick_label_skip {
        write!(writer, r#"<c:tickLblSkip val="{}"/>"#, skip)?;
    }
    if let Some(skip) = axis.tick_mark_skip {
        write!(writer, r#"<c:tickMarkSkip val="{}"/>"#, skip)?;
    }
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
    let mut entry_indexes = std::collections::HashSet::with_capacity(legend.entries.len());
    for entry in &legend.entries {
        if !entry_indexes.insert(entry.index) {
            return Err(invalid_chart_input(format!(
                "chart legend contains duplicate entry index {}",
                entry.index
            )));
        }
        write!(writer, "<c:legendEntry>")?;
        write!(writer, r#"<c:idx val="{}"/>"#, entry.index)?;
        write!(
            writer,
            r#"<c:delete val="{}"/>"#,
            if entry.deleted { "1" } else { "0" }
        )?;
        write!(writer, "</c:legendEntry>")?;
    }
    if let Some(layout) = legend.layout.as_ref() {
        write_layout(writer, Some(layout))?;
    }
    write!(
        writer,
        r#"<c:overlay val="{}"/>"#,
        if legend.overlay { "1" } else { "0" }
    )?;
    write!(writer, "</c:legend>")?;
    Ok(())
}
