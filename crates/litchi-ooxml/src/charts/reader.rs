//! Chart XML reader.
//!
//! This module provides functionality to parse chart XML files
//! from OOXML packages.

use crate::charts::axis::{Axis, CategoryAxis, DateAxis, SeriesAxis, ValueAxis};
use crate::charts::chart::{Chart, View3D, WallFloor};
use crate::charts::legend::Legend;
use crate::charts::models::{DataSourceRef, NumericData, RichText, StringData, TitleText};
use crate::charts::plot_area::{
    Area3DTypeGroup, AreaTypeGroup, Bar3DTypeGroup, BarShape, BarTypeGroup, BubbleTypeGroup,
    DoughnutTypeGroup, Line3DTypeGroup, LineTypeGroup, Pie3DTypeGroup, PieTypeGroup, PlotArea,
    RadarTypeGroup, ScatterTypeGroup, StockTypeGroup, Surface3DTypeGroup, SurfaceTypeGroup,
    TypeGroup, TypeGroupCommon,
};
use crate::charts::series::Series;
use crate::charts::types::{
    AxisPosition, BarDirection, BarGrouping, DisplayBlanks, LegendPosition, RadarStyle,
    ScatterStyle,
};
use crate::common::xml::decode_xml_reference;
use crate::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::io::BufRead;

/// Parse a chart XML document.
pub fn parse_chart<R: BufRead>(reader: R) -> Result<Chart> {
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(false);

    let mut chart = Chart::new();
    let mut buf = Vec::new();
    let mut saw_chart = false;
    let mut closed_chart = false;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"chart" => saw_chart = true,
                    b"chartSpace" => {},
                    b"title" => {
                        chart.title = Some(parse_title(&mut xml_reader)?);
                    },
                    b"autoTitleDeleted" => {
                        chart.auto_title_deleted = parse_bool_attr(e)?;
                    },
                    b"view3D" => {
                        chart.view_3d = Some(parse_view_3d(&mut xml_reader)?);
                    },
                    b"floor" => {
                        chart.floor = Some(parse_wall_floor(&mut xml_reader)?);
                    },
                    b"backWall" => {
                        chart.back_wall = Some(parse_wall_floor(&mut xml_reader)?);
                    },
                    b"sideWall" => {
                        chart.side_wall = Some(parse_wall_floor(&mut xml_reader)?);
                    },
                    b"plotArea" => {
                        chart.plot_area = parse_plot_area(&mut xml_reader)?;
                    },
                    b"legend" => {
                        chart.legend = Some(parse_legend(&mut xml_reader)?);
                    },
                    b"plotVisOnly" => {
                        chart.plot_visible_only = parse_bool_attr(e)?;
                    },
                    b"dispBlanksAs" => {
                        chart.display_blanks_as = parse_display_blanks(e)?;
                    },
                    b"date1904" => {
                        chart.date_1904 = parse_bool_attr(e)?;
                    },
                    b"roundedCorners" => {
                        chart.rounded_corners = parse_bool_attr(e)?;
                    },
                    b"style" => {
                        chart.style = parse_u32_attr(e, b"val");
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"chart" => {
                closed_chart = true;
            },
            Ok(Event::Eof) if saw_chart && closed_chart => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "chart XML has no complete chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(chart)
}

fn parse_title<R: BufRead>(reader: &mut Reader<R>) -> Result<TitleText> {
    let mut text = String::new();
    let mut buf = Vec::new();
    let mut in_text = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"t" => {
                in_text = true;
            },
            Ok(Event::Text(e)) if in_text => {
                text.push_str(&e.decode().map_err(|e| OoxmlError::Xml(e.to_string()))?);
            },
            Ok(Event::CData(e)) if in_text => {
                text.push_str(&e.decode().map_err(|e| OoxmlError::Xml(e.to_string()))?);
            },
            Ok(Event::GeneralRef(reference)) if in_text => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"t" => {
                in_text = false;
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"title" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(TitleText::Literal(RichText::new(text)))
}

fn parse_view_3d<R: BufRead>(reader: &mut Reader<R>) -> Result<View3D> {
    let mut view = View3D::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"rotX" => view.rot_x = parse_u32_attr(e, b"val"),
                    b"rotY" => view.rot_y = parse_u32_attr(e, b"val"),
                    b"perspective" => view.perspective = parse_u32_attr(e, b"val"),
                    b"hPercent" => view.height_percent = parse_u32_attr(e, b"val"),
                    b"depthPercent" => view.depth_percent = parse_u32_attr(e, b"val"),
                    b"rAngAx" => view.right_angle_axes = parse_bool_attr(e)?,
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"view3D" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(view)
}

fn parse_wall_floor<R: BufRead>(reader: &mut Reader<R>) -> Result<WallFloor> {
    let mut wall_floor = WallFloor::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"thickness" =>
            {
                wall_floor.thickness = parse_u32_attr(e, b"val");
            },
            Ok(Event::End(ref e)) => {
                let tag_name = e.local_name();
                if tag_name.as_ref() == b"floor"
                    || tag_name.as_ref() == b"backWall"
                    || tag_name.as_ref() == b"sideWall"
                {
                    break;
                }
            },
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(wall_floor)
}

fn parse_plot_area<R: BufRead>(reader: &mut Reader<R>) -> Result<PlotArea> {
    let mut plot_area = PlotArea::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"barChart" => {
                        if let Some(group) = parse_bar_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Bar(group));
                        }
                    },
                    b"bar3DChart" => {
                        if let Some(group) = parse_bar_3d_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Bar3D(group));
                        }
                    },
                    b"lineChart" => {
                        if let Some(group) = parse_line_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Line(group));
                        }
                    },
                    b"pieChart" => {
                        if let Some(group) = parse_pie_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Pie(group));
                        }
                    },
                    b"areaChart" => {
                        if let Some(group) = parse_area_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Area(group));
                        }
                    },
                    b"area3DChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Area3D(parse_area_3d_chart(reader)?));
                    },
                    b"bubbleChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Bubble(parse_bubble_chart(reader)?));
                    },
                    b"doughnutChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Doughnut(parse_doughnut_chart(reader)?));
                    },
                    b"line3DChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Line3D(parse_line_3d_chart(reader)?));
                    },
                    b"pie3DChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Pie3D(parse_pie_3d_chart(reader)?));
                    },
                    b"radarChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Radar(parse_radar_chart(reader)?));
                    },
                    b"scatterChart" => {
                        if let Some(group) = parse_scatter_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Scatter(group));
                        }
                    },
                    b"stockChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Stock(parse_stock_chart(reader)?));
                    },
                    b"surfaceChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Surface(parse_surface_chart(reader)?));
                    },
                    b"surface3DChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::Surface3D(parse_surface_3d_chart(reader)?));
                    },
                    b"catAx" => {
                        if let Some(axis) = parse_category_axis(reader)? {
                            plot_area.axes.push(Axis::Category(axis));
                        }
                    },
                    b"valAx" => {
                        if let Some(axis) = parse_value_axis(reader)? {
                            plot_area.axes.push(Axis::Value(axis));
                        }
                    },
                    b"dateAx" => {
                        if let Some(axis) = parse_date_axis(reader)? {
                            plot_area.axes.push(Axis::Date(axis));
                        }
                    },
                    b"serAx" => {
                        if let Some(axis) = parse_series_axis(reader)? {
                            plot_area.axes.push(Axis::Series(axis));
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"plotArea" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(plot_area)
}

fn parse_common_type_group<R: BufRead>(
    reader: &mut Reader<R>,
    end_name: &[u8],
    mut extra: impl FnMut(&BytesStart<'_>) -> Result<()>,
) -> Result<TypeGroupCommon> {
    let mut common = TypeGroupCommon::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"varyColors" => common.vary_colors = parse_bool_attr(element)?,
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    _ => extra(element)?,
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == end_name => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart type group".to_string(),
                ));
            },
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
        buf.clear();
    }
    Ok(common)
}

fn parse_area_3d_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Area3DTypeGroup> {
    let mut grouping = BarGrouping::Standard;
    let common = parse_common_type_group(reader, b"area3DChart", |element| {
        if element.local_name().as_ref() == b"grouping" {
            grouping = parse_grouping(element)?;
        }
        Ok(())
    })?;
    let mut group = Area3DTypeGroup::new(grouping);
    group.common = common;
    Ok(group)
}

fn parse_bubble_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<BubbleTypeGroup> {
    let mut bubble_3d = false;
    let mut bubble_scale = None;
    let mut show_negative_bubbles = true;
    let mut size_represents = "area".to_string();
    let common = parse_common_type_group(reader, b"bubbleChart", |element| {
        match element.local_name().as_ref() {
            b"bubble3D" => bubble_3d = parse_bool_attr(element)?,
            b"bubbleScale" => {
                let value = required_u32_attr(element, "bubble scale")?;
                if value > 300 {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "chart bubble scale {value} exceeds 300"
                    )));
                }
                bubble_scale = Some(value);
            },
            b"showNegBubbles" => show_negative_bubbles = parse_bool_attr(element)?,
            b"sizeRepresents" => {
                let value = get_attr(element, b"val")
                    .ok_or_else(|| missing_attribute("chart bubble size representation"))?;
                match value.as_slice() {
                    b"area" | b"w" => {
                        size_represents = String::from_utf8_lossy(&value).into_owned();
                    },
                    _ => {
                        return Err(invalid_attribute(
                            "chart bubble size representation",
                            &value,
                        ));
                    },
                }
            },
            _ => {},
        }
        Ok(())
    })?;
    let mut group = BubbleTypeGroup::new();
    group.common = common;
    group.bubble_3d = bubble_3d;
    group.bubble_scale = bubble_scale;
    group.show_negative_bubbles = show_negative_bubbles;
    group.size_represents = size_represents;
    Ok(group)
}

fn parse_doughnut_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<DoughnutTypeGroup> {
    let mut first_slice_angle = 0;
    let mut hole_size = 50;
    let common = parse_common_type_group(reader, b"doughnutChart", |element| {
        match element.local_name().as_ref() {
            b"firstSliceAng" => {
                first_slice_angle = required_u32_attr(element, "first-slice angle")?;
                if first_slice_angle > 360 {
                    return Err(OoxmlError::InvalidFormat(
                        "chart first-slice angle exceeds 360".to_string(),
                    ));
                }
            },
            b"holeSize" => {
                hole_size = required_u32_attr(element, "doughnut hole size")?;
                if !(10..=90).contains(&hole_size) {
                    return Err(OoxmlError::InvalidFormat(
                        "chart doughnut hole size must be between 10 and 90".to_string(),
                    ));
                }
            },
            _ => {},
        }
        Ok(())
    })?;
    let mut group = DoughnutTypeGroup::new();
    group.common = common;
    group.first_slice_angle = first_slice_angle;
    group.hole_size = hole_size;
    Ok(group)
}

fn parse_line_3d_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Line3DTypeGroup> {
    let mut grouping = BarGrouping::Standard;
    let common = parse_common_type_group(reader, b"line3DChart", |element| {
        if element.local_name().as_ref() == b"grouping" {
            grouping = parse_grouping(element)?;
        }
        Ok(())
    })?;
    let mut group = Line3DTypeGroup::new(grouping);
    group.common = common;
    Ok(group)
}

fn parse_pie_3d_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Pie3DTypeGroup> {
    let mut group = Pie3DTypeGroup::new();
    group.common = parse_common_type_group(reader, b"pie3DChart", |_| Ok(()))?;
    Ok(group)
}

fn parse_radar_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<RadarTypeGroup> {
    let mut style = RadarStyle::Standard;
    let common = parse_common_type_group(reader, b"radarChart", |element| {
        if element.local_name().as_ref() == b"radarStyle" {
            let value =
                get_attr(element, b"val").ok_or_else(|| missing_attribute("chart radar style"))?;
            style = match value.as_slice() {
                b"standard" => RadarStyle::Standard,
                b"filled" => RadarStyle::Filled,
                b"marker" => RadarStyle::Marker,
                _ => return Err(invalid_attribute("chart radar style", &value)),
            };
        }
        Ok(())
    })?;
    let mut group = RadarTypeGroup::new(style);
    group.common = common;
    Ok(group)
}

fn parse_stock_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<StockTypeGroup> {
    let mut group = StockTypeGroup::new();
    group.common = parse_common_type_group(reader, b"stockChart", |_| Ok(()))?;
    Ok(group)
}

fn parse_surface_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<SurfaceTypeGroup> {
    let mut wireframe = false;
    let common = parse_common_type_group(reader, b"surfaceChart", |element| {
        if element.local_name().as_ref() == b"wireframe" {
            wireframe = parse_bool_attr(element)?;
        }
        Ok(())
    })?;
    let mut group = SurfaceTypeGroup::new();
    group.common = common;
    group.wireframe = wireframe;
    Ok(group)
}

fn parse_surface_3d_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Surface3DTypeGroup> {
    let mut wireframe = false;
    let common = parse_common_type_group(reader, b"surface3DChart", |element| {
        if element.local_name().as_ref() == b"wireframe" {
            wireframe = parse_bool_attr(element)?;
        }
        Ok(())
    })?;
    let mut group = Surface3DTypeGroup::new();
    group.common = common;
    group.wireframe = wireframe;
    Ok(group)
}

fn parse_bar_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<BarTypeGroup>> {
    let mut direction = BarDirection::Column;
    let mut grouping = BarGrouping::Clustered;
    let mut common = TypeGroupCommon::new();
    let mut gap_width = None;
    let mut overlap = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"barDir" => {
                        if let Some(val) = get_attr(e, b"val") {
                            direction = if val.as_slice() == b"bar" {
                                BarDirection::Bar
                            } else if val.as_slice() == b"col" {
                                BarDirection::Column
                            } else {
                                return Err(invalid_attribute("chart bar direction", &val));
                            };
                        }
                    },
                    b"grouping" => {
                        grouping = parse_grouping(e)?;
                    },
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"gapWidth" => {
                        let value = required_u32_attr(e, "chart gap width")?;
                        if value > 500 {
                            return Err(OoxmlError::InvalidFormat(
                                "chart gap width exceeds 500".to_string(),
                            ));
                        }
                        gap_width = Some(value);
                    },
                    b"overlap" => {
                        let value = required_i32_attr(e, "chart overlap")?;
                        if !(-100..=100).contains(&value) {
                            return Err(OoxmlError::InvalidFormat(
                                "chart overlap must be between -100 and 100".to_string(),
                            ));
                        }
                        overlap = Some(value);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"barChart" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    let mut group = BarTypeGroup::new(direction, grouping);
    group.common = common;
    group.gap_width = gap_width;
    group.overlap = overlap;
    Ok(Some(group))
}

fn parse_bar_3d_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<Bar3DTypeGroup>> {
    let mut direction = BarDirection::Column;
    let mut grouping = BarGrouping::Clustered;
    let mut common = TypeGroupCommon::new();
    let mut gap_width = None;
    let mut gap_depth = None;
    let mut shape = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"barDir" => {
                        if let Some(val) = get_attr(e, b"val") {
                            direction = if val.as_slice() == b"bar" {
                                BarDirection::Bar
                            } else if val.as_slice() == b"col" {
                                BarDirection::Column
                            } else {
                                return Err(invalid_attribute("chart bar direction", &val));
                            };
                        }
                    },
                    b"grouping" => {
                        grouping = parse_grouping(e)?;
                    },
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"gapWidth" => {
                        gap_width = Some(bounded_u32_attr(e, "chart gap width", 0, 500)?);
                    },
                    b"gapDepth" => {
                        gap_depth = Some(bounded_u32_attr(e, "chart gap depth", 0, 500)?);
                    },
                    b"shape" => {
                        let value = get_attr(e, b"val")
                            .ok_or_else(|| missing_attribute("chart bar shape"))?;
                        shape = Some(match value.as_slice() {
                            b"box" => BarShape::Box,
                            b"cone" => BarShape::Cone,
                            b"coneToMax" => BarShape::ConeToMax,
                            b"cylinder" => BarShape::Cylinder,
                            b"pyramid" => BarShape::Pyramid,
                            b"pyramidToMax" => BarShape::PyramidToMax,
                            _ => return Err(invalid_attribute("chart bar shape", &value)),
                        });
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"bar3DChart" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    let mut group = Bar3DTypeGroup::new(direction, grouping);
    group.common = common;
    group.gap_width = gap_width;
    group.gap_depth = gap_depth;
    group.shape = shape;
    Ok(Some(group))
}

fn parse_line_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<LineTypeGroup>> {
    let mut grouping = BarGrouping::Standard;
    let mut common = TypeGroupCommon::new();
    let mut marker = true;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"grouping" => {
                        grouping = parse_grouping(e)?;
                    },
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"marker" => marker = parse_bool_attr(e)?,
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"lineChart" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    let mut group = LineTypeGroup::new(grouping);
    group.common = common;
    group.marker = marker;
    Ok(Some(group))
}

fn parse_pie_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<PieTypeGroup>> {
    let mut common = TypeGroupCommon::new();
    let mut first_slice_angle = 0;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"firstSliceAng" => {
                        first_slice_angle = bounded_u32_attr(e, "chart first-slice angle", 0, 360)?;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"pieChart" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    let mut group = PieTypeGroup::new();
    group.common = common;
    group.first_slice_angle = first_slice_angle;
    Ok(Some(group))
}

fn parse_area_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<AreaTypeGroup>> {
    let mut grouping = BarGrouping::Standard;
    let mut common = TypeGroupCommon::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"grouping" => {
                        grouping = parse_grouping(e)?;
                    },
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"areaChart" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    let mut group = AreaTypeGroup::new(grouping);
    group.common = common;
    Ok(Some(group))
}

fn parse_scatter_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<ScatterTypeGroup>> {
    let mut style = ScatterStyle::LineMarker;
    let mut common = TypeGroupCommon::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"scatterStyle" => {
                        if let Some(val) = get_attr(e, b"val") {
                            style = match val.as_slice() {
                                b"line" => ScatterStyle::Line,
                                b"marker" => ScatterStyle::Marker,
                                b"none" => ScatterStyle::None,
                                b"smooth" => ScatterStyle::Smooth,
                                b"smoothMarker" => ScatterStyle::SmoothMarker,
                                b"lineMarker" => ScatterStyle::LineMarker,
                                _ => {
                                    return Err(invalid_attribute("chart scatter style", &val));
                                },
                            };
                        }
                    },
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"scatterChart" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    let mut group = ScatterTypeGroup::new(style);
    group.common = common;
    Ok(Some(group))
}

fn parse_series<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<Series>> {
    let mut series = Series::new(0);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"idx" => {
                        series.index = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"order" => {
                        series.order = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"tx" => {
                        series.title = parse_series_title(reader)?;
                    },
                    b"cat" => {
                        series.categories = parse_string_data(reader)?;
                    },
                    b"val" => {
                        series.values = parse_numeric_data(reader)?;
                    },
                    b"xVal" => {
                        series.x_values = parse_numeric_data(reader)?;
                    },
                    b"yVal" => {
                        series.y_values = parse_numeric_data(reader)?;
                    },
                    b"bubbleSize" => {
                        series.bubble_sizes = parse_numeric_data(reader)?;
                    },
                    b"explosion" => {
                        series.explosion = Some(required_u32_attr(e, "series explosion")?);
                    },
                    b"bubble3D" => {
                        series.bubble_3d = parse_bool_attr(e)?;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"ser" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(series))
}

fn parse_string_data<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<StringData>> {
    let mut data = StringData::from_values(Vec::new());
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"f" => {
                let formula = parse_text_element(reader, b"f")?;
                data.source_ref = Some(DataSourceRef::new(formula));
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pt" => {
                if let Some(text) = parse_point_text(reader)? {
                    data.values.push(text);
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cat" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(data))
}

fn parse_numeric_data<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<NumericData>> {
    let mut data = NumericData::from_values(Vec::new());
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"f" => {
                let formula = parse_text_element(reader, b"f")?;
                data.source_ref = Some(DataSourceRef::new(formula));
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"formatCode" => {
                data.format_code = Some(parse_text_element(reader, b"formatCode")?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pt" => {
                if let Some(val) = parse_point_value(reader)? {
                    data.values.push(val);
                }
            },
            Ok(Event::End(ref e))
                if matches!(
                    e.local_name().as_ref(),
                    b"val" | b"xVal" | b"yVal" | b"bubbleSize"
                ) =>
            {
                break;
            },
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(data))
}

fn parse_series_title<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<TitleText>> {
    let mut title = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"f" => {
                let formula = parse_text_element(reader, b"f")?;
                set_title(
                    &mut title,
                    TitleText::Reference(DataSourceRef::new(formula)),
                )?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"v" => {
                let value = parse_text_element(reader, b"v")?;
                set_title(&mut title, TitleText::Literal(RichText::new(value)))?;
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"tx" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart series title".to_string(),
                ));
            },
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
        buf.clear();
    }
    Ok(title)
}

fn set_title(target: &mut Option<TitleText>, title: TitleText) -> Result<()> {
    if target.replace(title).is_some() {
        return Err(OoxmlError::InvalidFormat(
            "chart series has duplicate title values".to_string(),
        ));
    }
    Ok(())
}

fn parse_text_element<R: BufRead>(reader: &mut Reader<R>, end_name: &[u8]) -> Result<String> {
    let mut text = String::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Text(value)) => {
                text.push_str(
                    &value
                        .decode()
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?,
                );
            },
            Ok(Event::CData(value)) => {
                text.push_str(
                    &value
                        .decode()
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?,
                );
            },
            Ok(Event::GeneralRef(reference)) => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == end_name => break,
            Ok(Event::Start(_)) | Ok(Event::Empty(_)) => {
                return Err(OoxmlError::InvalidFormat(
                    "chart text element contains nested markup".to_string(),
                ));
            },
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart text element".to_string(),
                ));
            },
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
        buf.clear();
    }
    Ok(text)
}

fn parse_point_text<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<String>> {
    let mut text = String::new();
    let mut buf = Vec::new();
    let mut in_v = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"v" => {
                in_v = true;
            },
            Ok(Event::Text(e)) if in_v => {
                text.push_str(&e.decode().map_err(|e| OoxmlError::Xml(e.to_string()))?);
            },
            Ok(Event::CData(e)) if in_v => {
                text.push_str(&e.decode().map_err(|e| OoxmlError::Xml(e.to_string()))?);
            },
            Ok(Event::GeneralRef(reference)) if in_v => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"v" => {
                in_v = false;
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"pt" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(text))
}

fn parse_point_value<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<f64>> {
    if let Some(text) = parse_point_text(reader)? {
        Ok(Some(text.trim().parse::<f64>().map_err(|_| {
            OoxmlError::InvalidFormat(format!("invalid chart numeric point '{text}'"))
        })?))
    } else {
        Ok(None)
    }
}

fn parse_category_axis<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<CategoryAxis>> {
    let mut axis_id = 0;
    let mut cross_axis_id = 0;
    let mut position = AxisPosition::Bottom;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"axId" => {
                        axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"crossAx" => {
                        cross_axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"axPos" => {
                        position = parse_axis_position(e)?;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"catAx" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(CategoryAxis::new(axis_id, position, cross_axis_id)))
}

fn parse_value_axis<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<ValueAxis>> {
    let mut axis_id = 0;
    let mut cross_axis_id = 0;
    let mut position = AxisPosition::Left;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"axId" => {
                        axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"crossAx" => {
                        cross_axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"axPos" => {
                        position = parse_axis_position(e)?;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"valAx" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(ValueAxis::new(axis_id, position, cross_axis_id)))
}

fn parse_date_axis<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<DateAxis>> {
    let mut axis_id = 0;
    let mut cross_axis_id = 0;
    let mut position = AxisPosition::Bottom;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"axId" => {
                        axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"crossAx" => {
                        cross_axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"axPos" => {
                        position = parse_axis_position(e)?;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dateAx" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(DateAxis::new(axis_id, position, cross_axis_id)))
}

fn parse_series_axis<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<SeriesAxis>> {
    let mut axis_id = 0;
    let mut cross_axis_id = 0;
    let mut position = AxisPosition::Bottom;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"axId" => {
                        axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"crossAx" => {
                        cross_axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"axPos" => {
                        position = parse_axis_position(e)?;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"serAx" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(SeriesAxis::new(axis_id, position, cross_axis_id)))
}

fn parse_legend<R: BufRead>(reader: &mut Reader<R>) -> Result<Legend> {
    let mut position = LegendPosition::Right;
    let mut overlay = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"legendPos" => {
                        if let Some(val) = get_attr(e, b"val") {
                            position = match val.as_slice() {
                                b"b" => LegendPosition::Bottom,
                                b"l" => LegendPosition::Left,
                                b"r" => LegendPosition::Right,
                                b"t" => LegendPosition::Top,
                                b"tr" => LegendPosition::TopRight,
                                _ => LegendPosition::Right,
                            };
                        }
                    },
                    b"overlay" => {
                        overlay = parse_bool_attr(e)?;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"legend" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(Legend::new(position).with_overlay(overlay))
}

#[inline]
fn parse_grouping(e: &BytesStart) -> Result<BarGrouping> {
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
fn parse_axis_position(e: &BytesStart) -> Result<AxisPosition> {
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

#[inline]
fn parse_display_blanks(e: &BytesStart) -> crate::error::Result<DisplayBlanks> {
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
fn parse_bool_attr(e: &BytesStart) -> crate::error::Result<bool> {
    if let Some(val) = get_attr(e, b"val") {
        match val.as_slice() {
            b"1" | b"true" => Ok(true),
            b"0" | b"false" => Ok(false),
            _ => Err(invalid_attribute("chart boolean", &val)),
        }
    } else {
        Ok(true)
    }
}

fn invalid_attribute(description: &str, value: &[u8]) -> OoxmlError {
    OoxmlError::InvalidFormat(format!(
        "invalid {description} '{}'",
        String::from_utf8_lossy(value)
    ))
}

#[inline]
fn parse_u32_attr(e: &BytesStart, attr_name: &[u8]) -> Option<u32> {
    get_attr(e, attr_name).and_then(|v| std::str::from_utf8(&v).ok()?.parse().ok())
}

fn required_u32_attr(element: &BytesStart<'_>, description: &str) -> Result<u32> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute(description))?;
    std::str::from_utf8(&value)
        .ok()
        .and_then(|value| value.parse().ok())
        .ok_or_else(|| invalid_attribute(description, &value))
}

fn required_i32_attr(element: &BytesStart<'_>, description: &str) -> Result<i32> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute(description))?;
    std::str::from_utf8(&value)
        .ok()
        .and_then(|value| value.parse().ok())
        .ok_or_else(|| invalid_attribute(description, &value))
}

fn bounded_u32_attr(
    element: &BytesStart<'_>,
    description: &str,
    minimum: u32,
    maximum: u32,
) -> Result<u32> {
    let value = required_u32_attr(element, description)?;
    if !(minimum..=maximum).contains(&value) {
        return Err(OoxmlError::InvalidFormat(format!(
            "{description} must be between {minimum} and {maximum}"
        )));
    }
    Ok(value)
}

fn missing_attribute(description: &str) -> OoxmlError {
    OoxmlError::InvalidFormat(format!("{description} is missing its value"))
}

#[inline]
fn get_attr(e: &BytesStart, name: &[u8]) -> Option<Vec<u8>> {
    e.attributes()
        .filter_map(|a| a.ok())
        .find(|a| a.key.as_ref() == name)
        .map(|a| a.value.to_vec())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_prefixed_chart_content() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue &amp; <![CDATA[Growth]]></a:t></a:r></a:p>
                </c:rich></c:tx></c:title><c:plotArea><c:barChart>
                <c:barDir val="bar"/><c:grouping val="stacked"/><c:ser>
                    <c:idx val="2"/><c:order val="1"/>
                    <c:cat><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:pt idx="0"><c:v>East</c:v></c:pt>
                        </c:strCache></c:strRef></c:cat>
                    <c:val><c:numRef><c:f>Sheet1!$B$1</c:f><c:numCache><c:formatCode>0.0</c:formatCode><c:pt idx="0"><c:v>42.5</c:v></c:pt>
                        </c:numCache></c:numRef></c:val>
                </c:ser></c:barChart></c:plotArea>
                <c:legend><c:legendPos val="b"/><c:overlay val="1"/></c:legend>
            </c:chart></c:chartSpace>"#;

        let chart = parse_chart(xml.as_slice()).unwrap();
        let Some(TitleText::Literal(title)) = chart.title.as_ref() else {
            panic!("expected a literal chart title");
        };
        assert_eq!(title.text, "Revenue & Growth");
        assert_eq!(chart.plot_area.type_groups.len(), 1);
        let TypeGroup::Bar(group) = &chart.plot_area.type_groups[0] else {
            panic!("expected a bar chart");
        };
        assert_eq!(group.direction, BarDirection::Bar);
        assert_eq!(group.grouping, BarGrouping::Stacked);
        assert_eq!(group.common.series.len(), 1);
        assert_eq!(group.common.series[0].index, 2);
        assert_eq!(
            group.common.series[0].categories.as_ref().unwrap().values,
            ["East"]
        );
        assert_eq!(
            group.common.series[0]
                .categories
                .as_ref()
                .unwrap()
                .source_ref
                .as_ref()
                .unwrap()
                .formula,
            "Sheet1!$A$1"
        );
        assert_eq!(
            group.common.series[0].values.as_ref().unwrap().values,
            [42.5]
        );
        let values = group.common.series[0].values.as_ref().unwrap();
        assert_eq!(values.source_ref.as_ref().unwrap().formula, "Sheet1!$B$1");
        assert_eq!(values.format_code.as_deref(), Some("0.0"));
        assert_eq!(chart.legend.unwrap().position, LegendPosition::Bottom);
    }

    #[test]
    fn rejects_truncated_and_invalid_chart_values() {
        for xml in [
            br#"<c:chartSpace xmlns:c="urn:test"><c:chart><c:plotArea>"#.as_slice(),
            br#"<c:chartSpace xmlns:c="urn:test"><c:chart><c:plotVisOnly val="yes"/></c:chart></c:chartSpace>"#.as_slice(),
            br#"<c:chartSpace xmlns:c="urn:test"><c:chart><c:dispBlanksAs val="empty"/></c:chart></c:chartSpace>"#.as_slice(),
        ] {
            assert!(parse_chart(xml).is_err());
        }
    }

    #[test]
    fn writer_round_trips_every_modeled_chart_group() {
        let mut doughnut = DoughnutTypeGroup::new();
        doughnut.first_slice_angle = 45;
        doughnut.hole_size = 60;
        let mut surface = SurfaceTypeGroup::new();
        surface.wireframe = true;
        let mut surface_3d = Surface3DTypeGroup::new();
        surface_3d.wireframe = true;
        let mut bubble = BubbleTypeGroup::new();
        bubble.bubble_scale = Some(125);
        bubble.show_negative_bubbles = false;

        let mut chart = Chart::new();
        chart.plot_area.type_groups = vec![
            TypeGroup::Area(AreaTypeGroup::new(BarGrouping::Standard)),
            TypeGroup::Area3D(Area3DTypeGroup::new(BarGrouping::Stacked)),
            TypeGroup::Bar(BarTypeGroup::new(
                BarDirection::Column,
                BarGrouping::Clustered,
            )),
            TypeGroup::Bar3D(Bar3DTypeGroup::new(BarDirection::Bar, BarGrouping::Stacked)),
            TypeGroup::Bubble(bubble),
            TypeGroup::Doughnut(doughnut),
            TypeGroup::Line(LineTypeGroup::new(BarGrouping::Standard)),
            TypeGroup::Line3D(Line3DTypeGroup::new(BarGrouping::PercentStacked)),
            TypeGroup::Pie(PieTypeGroup::new()),
            TypeGroup::Pie3D(Pie3DTypeGroup::new()),
            TypeGroup::Radar(RadarTypeGroup::new(RadarStyle::Filled)),
            TypeGroup::Scatter(ScatterTypeGroup::new(ScatterStyle::Smooth)),
            TypeGroup::Stock(StockTypeGroup::new()),
            TypeGroup::Surface(surface),
            TypeGroup::Surface3D(surface_3d),
        ];

        let mut xml = Vec::new();
        crate::charts::writer::write_chart(&mut xml, &chart).unwrap();
        let parsed = parse_chart(xml.as_slice()).unwrap();

        assert_eq!(parsed.plot_area.type_groups.len(), 15);
        assert!(matches!(
            parsed.plot_area.type_groups[0],
            TypeGroup::Area(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[1],
            TypeGroup::Area3D(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[4],
            TypeGroup::Bubble(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[5],
            TypeGroup::Doughnut(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[7],
            TypeGroup::Line3D(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[9],
            TypeGroup::Pie3D(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[10],
            TypeGroup::Radar(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[12],
            TypeGroup::Stock(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[13],
            TypeGroup::Surface(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[14],
            TypeGroup::Surface3D(_)
        ));
        let TypeGroup::Doughnut(group) = &parsed.plot_area.type_groups[5] else {
            unreachable!();
        };
        assert_eq!(group.first_slice_angle, 45);
        assert_eq!(group.hole_size, 60);
        let TypeGroup::Bubble(group) = &parsed.plot_area.type_groups[4] else {
            unreachable!();
        };
        assert_eq!(group.bubble_scale, Some(125));
        assert!(!group.show_negative_bubbles);
    }
}
