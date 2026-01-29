//! Chart XML reader.
//!
//! This module provides functionality to parse chart XML files
//! from OOXML packages.

use crate::ooxml::charts::axis::{Axis, CategoryAxis, DateAxis, SeriesAxis, ValueAxis};
use crate::ooxml::charts::chart::{Chart, View3D, WallFloor};
use crate::ooxml::charts::legend::Legend;
use crate::ooxml::charts::models::{NumericData, RichText, StringData, TitleText};
use crate::ooxml::charts::plot_area::{
    AreaTypeGroup, Bar3DTypeGroup, BarTypeGroup, LineTypeGroup, PieTypeGroup, PlotArea,
    ScatterTypeGroup, TypeGroup, TypeGroupCommon,
};
use crate::ooxml::charts::series::Series;
use crate::ooxml::charts::types::{
    AxisPosition, BarDirection, BarGrouping, DisplayBlanks, LegendPosition, ScatterStyle,
};
use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::io::BufRead;

/// Parse a chart XML document.
pub fn parse_chart<R: BufRead>(reader: R) -> Result<Chart> {
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut chart = Chart::new();
    let mut buf = Vec::new();

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"c:chart" => {
                        parse_chart_element(&mut xml_reader, &mut chart)?;
                    },
                    b"c:title" => {
                        chart.title = Some(parse_title(&mut xml_reader)?);
                    },
                    b"c:autoTitleDeleted" => {
                        chart.auto_title_deleted = parse_bool_attr(e)?;
                    },
                    b"c:view3D" => {
                        chart.view_3d = Some(parse_view_3d(&mut xml_reader)?);
                    },
                    b"c:floor" => {
                        chart.floor = Some(parse_wall_floor(&mut xml_reader)?);
                    },
                    b"c:backWall" => {
                        chart.back_wall = Some(parse_wall_floor(&mut xml_reader)?);
                    },
                    b"c:sideWall" => {
                        chart.side_wall = Some(parse_wall_floor(&mut xml_reader)?);
                    },
                    b"c:plotArea" => {
                        chart.plot_area = parse_plot_area(&mut xml_reader)?;
                    },
                    b"c:legend" => {
                        chart.legend = Some(parse_legend(&mut xml_reader)?);
                    },
                    b"c:plotVisOnly" => {
                        chart.plot_visible_only = parse_bool_attr(e)?;
                    },
                    b"c:dispBlanksAs" => {
                        chart.display_blanks_as = parse_display_blanks(e)?;
                    },
                    b"c:date1904" => {
                        chart.date_1904 = parse_bool_attr(e)?;
                    },
                    b"c:roundedCorners" => {
                        chart.rounded_corners = parse_bool_attr(e)?;
                    },
                    b"c:style" => {
                        chart.style = parse_u32_attr(e, b"val");
                    },
                    _ => {},
                }
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(chart)
}

fn parse_chart_element<R: BufRead>(reader: &mut Reader<R>, _chart: &mut Chart) -> Result<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:chart" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }
    Ok(())
}

fn parse_title<R: BufRead>(reader: &mut Reader<R>) -> Result<TitleText> {
    let mut text = String::new();
    let mut buf = Vec::new();
    let mut in_text = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"a:t" => {
                in_text = true;
            },
            Ok(Event::Text(e)) if in_text => {
                text.push_str(
                    std::str::from_utf8(e.as_ref()).map_err(|e| OoxmlError::Xml(e.to_string()))?,
                );
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"a:t" => {
                in_text = false;
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:title" => break,
            Ok(Event::Eof) => break,
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
                    b"c:rotX" => view.rot_x = parse_u32_attr(e, b"val"),
                    b"c:rotY" => view.rot_y = parse_u32_attr(e, b"val"),
                    b"c:perspective" => view.perspective = parse_u32_attr(e, b"val"),
                    b"c:hPercent" => view.height_percent = parse_u32_attr(e, b"val"),
                    b"c:depthPercent" => view.depth_percent = parse_u32_attr(e, b"val"),
                    b"c:rAngAx" => view.right_angle_axes = parse_bool_attr(e)?,
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:view3D" => break,
            Ok(Event::Eof) => break,
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
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"c:thickness" {
                    wall_floor.thickness = parse_u32_attr(e, b"val");
                }
            },
            Ok(Event::End(ref e)) => {
                let tag_name = e.local_name();
                if tag_name.as_ref() == b"c:floor"
                    || tag_name.as_ref() == b"c:backWall"
                    || tag_name.as_ref() == b"c:sideWall"
                {
                    break;
                }
            },
            Ok(Event::Eof) => break,
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
                    b"c:barChart" => {
                        if let Some(group) = parse_bar_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Bar(group));
                        }
                    },
                    b"c:bar3DChart" => {
                        if let Some(group) = parse_bar_3d_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Bar3D(group));
                        }
                    },
                    b"c:lineChart" => {
                        if let Some(group) = parse_line_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Line(group));
                        }
                    },
                    b"c:pieChart" => {
                        if let Some(group) = parse_pie_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Pie(group));
                        }
                    },
                    b"c:areaChart" => {
                        if let Some(group) = parse_area_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Area(group));
                        }
                    },
                    b"c:scatterChart" => {
                        if let Some(group) = parse_scatter_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Scatter(group));
                        }
                    },
                    b"c:catAx" => {
                        if let Some(axis) = parse_category_axis(reader)? {
                            plot_area.axes.push(Axis::Category(axis));
                        }
                    },
                    b"c:valAx" => {
                        if let Some(axis) = parse_value_axis(reader)? {
                            plot_area.axes.push(Axis::Value(axis));
                        }
                    },
                    b"c:dateAx" => {
                        if let Some(axis) = parse_date_axis(reader)? {
                            plot_area.axes.push(Axis::Date(axis));
                        }
                    },
                    b"c:serAx" => {
                        if let Some(axis) = parse_series_axis(reader)? {
                            plot_area.axes.push(Axis::Series(axis));
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:plotArea" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(plot_area)
}

fn parse_bar_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<BarTypeGroup>> {
    let mut direction = BarDirection::Column;
    let mut grouping = BarGrouping::Clustered;
    let mut common = TypeGroupCommon::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"c:barDir" => {
                        if let Some(val) = get_attr(e, b"val") {
                            direction = if val.as_slice() == b"bar" {
                                BarDirection::Bar
                            } else {
                                BarDirection::Column
                            };
                        }
                    },
                    b"c:grouping" => {
                        grouping = parse_grouping(e);
                    },
                    b"c:varyColors" => {
                        common.vary_colors = parse_bool_attr(e).unwrap_or(false);
                    },
                    b"c:ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:barChart" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    let mut group = BarTypeGroup::new(direction, grouping);
    group.common = common;
    Ok(Some(group))
}

fn parse_bar_3d_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<Bar3DTypeGroup>> {
    let mut direction = BarDirection::Column;
    let mut grouping = BarGrouping::Clustered;
    let mut common = TypeGroupCommon::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"c:barDir" => {
                        if let Some(val) = get_attr(e, b"val") {
                            direction = if val.as_slice() == b"bar" {
                                BarDirection::Bar
                            } else {
                                BarDirection::Column
                            };
                        }
                    },
                    b"c:grouping" => {
                        grouping = parse_grouping(e);
                    },
                    b"c:varyColors" => {
                        common.vary_colors = parse_bool_attr(e).unwrap_or(false);
                    },
                    b"c:ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:bar3DChart" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    let mut group = Bar3DTypeGroup::new(direction, grouping);
    group.common = common;
    Ok(Some(group))
}

fn parse_line_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<LineTypeGroup>> {
    let mut grouping = BarGrouping::Standard;
    let mut common = TypeGroupCommon::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"c:grouping" => {
                        grouping = parse_grouping(e);
                    },
                    b"c:varyColors" => {
                        common.vary_colors = parse_bool_attr(e).unwrap_or(false);
                    },
                    b"c:ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:lineChart" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    let mut group = LineTypeGroup::new(grouping);
    group.common = common;
    Ok(Some(group))
}

fn parse_pie_chart<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<PieTypeGroup>> {
    let mut common = TypeGroupCommon::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"c:varyColors" => {
                        common.vary_colors = parse_bool_attr(e).unwrap_or(true);
                    },
                    b"c:ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:pieChart" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    let mut group = PieTypeGroup::new();
    group.common = common;
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
                    b"c:grouping" => {
                        grouping = parse_grouping(e);
                    },
                    b"c:varyColors" => {
                        common.vary_colors = parse_bool_attr(e).unwrap_or(false);
                    },
                    b"c:ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:areaChart" => break,
            Ok(Event::Eof) => break,
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
                    b"c:scatterStyle" => {
                        if let Some(val) = get_attr(e, b"val") {
                            style = match val.as_slice() {
                                b"line" => ScatterStyle::Line,
                                b"marker" => ScatterStyle::Marker,
                                b"none" => ScatterStyle::None,
                                b"smooth" => ScatterStyle::Smooth,
                                b"smoothMarker" => ScatterStyle::SmoothMarker,
                                _ => ScatterStyle::LineMarker,
                            };
                        }
                    },
                    b"c:varyColors" => {
                        common.vary_colors = parse_bool_attr(e).unwrap_or(false);
                    },
                    b"c:ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:scatterChart" => break,
            Ok(Event::Eof) => break,
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
                    b"c:idx" => {
                        series.index = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"c:order" => {
                        series.order = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"c:cat" => {
                        series.categories = parse_string_data(reader)?;
                    },
                    b"c:val" => {
                        series.values = parse_numeric_data(reader)?;
                    },
                    b"c:xVal" => {
                        series.x_values = parse_numeric_data(reader)?;
                    },
                    b"c:yVal" => {
                        series.y_values = parse_numeric_data(reader)?;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:ser" => break,
            Ok(Event::Eof) => break,
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
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"c:pt" => {
                if let Some(text) = parse_point_text(reader)? {
                    data.values.push(text);
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:cat" => break,
            Ok(Event::Eof) => break,
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
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"c:pt" => {
                if let Some(val) = parse_point_value(reader)? {
                    data.values.push(val);
                }
            },
            Ok(Event::End(ref e))
                if matches!(e.local_name().as_ref(), b"c:val" | b"c:xVal" | b"c:yVal") =>
            {
                break;
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(data))
}

fn parse_point_text<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<String>> {
    let mut text = String::new();
    let mut buf = Vec::new();
    let mut in_v = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"c:v" => {
                in_v = true;
            },
            Ok(Event::Text(e)) if in_v => {
                text = std::str::from_utf8(e.as_ref())
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?
                    .to_string();
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:pt" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(text))
}

fn parse_point_value<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<f64>> {
    if let Some(text) = parse_point_text(reader)? {
        Ok(text.parse::<f64>().ok())
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
                    b"c:axId" => {
                        axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"c:crossAx" => {
                        cross_axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"c:axPos" => {
                        position = parse_axis_position(e);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:catAx" => break,
            Ok(Event::Eof) => break,
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
                    b"c:axId" => {
                        axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"c:crossAx" => {
                        cross_axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"c:axPos" => {
                        position = parse_axis_position(e);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:valAx" => break,
            Ok(Event::Eof) => break,
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
                    b"c:axId" => {
                        axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"c:crossAx" => {
                        cross_axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"c:axPos" => {
                        position = parse_axis_position(e);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:dateAx" => break,
            Ok(Event::Eof) => break,
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
                    b"c:axId" => {
                        axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"c:crossAx" => {
                        cross_axis_id = parse_u32_attr(e, b"val").unwrap_or(0);
                    },
                    b"c:axPos" => {
                        position = parse_axis_position(e);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:serAx" => break,
            Ok(Event::Eof) => break,
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
                    b"c:legendPos" => {
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
                    b"c:overlay" => {
                        overlay = parse_bool_attr(e).unwrap_or(false);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c:legend" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
        buf.clear();
    }

    Ok(Legend::new(position).with_overlay(overlay))
}

#[inline]
fn parse_grouping(e: &BytesStart) -> BarGrouping {
    if let Some(val) = get_attr(e, b"val") {
        match val.as_slice() {
            b"clustered" => BarGrouping::Clustered,
            b"stacked" => BarGrouping::Stacked,
            b"percentStacked" => BarGrouping::PercentStacked,
            _ => BarGrouping::Standard,
        }
    } else {
        BarGrouping::Standard
    }
}

#[inline]
fn parse_axis_position(e: &BytesStart) -> AxisPosition {
    if let Some(val) = get_attr(e, b"val") {
        match val.as_slice() {
            b"b" => AxisPosition::Bottom,
            b"l" => AxisPosition::Left,
            b"r" => AxisPosition::Right,
            b"t" => AxisPosition::Top,
            _ => AxisPosition::Bottom,
        }
    } else {
        AxisPosition::Bottom
    }
}

#[inline]
fn parse_display_blanks(e: &BytesStart) -> crate::ooxml::error::Result<DisplayBlanks> {
    if let Some(val) = get_attr(e, b"val") {
        Ok(match val.as_slice() {
            b"gap" => DisplayBlanks::Gap,
            b"span" => DisplayBlanks::Span,
            b"zero" => DisplayBlanks::Zero,
            _ => DisplayBlanks::Gap,
        })
    } else {
        Ok(DisplayBlanks::Gap)
    }
}

#[inline]
fn parse_bool_attr(e: &BytesStart) -> crate::ooxml::error::Result<bool> {
    if let Some(val) = get_attr(e, b"val") {
        Ok(val.as_slice() == b"1" || val.as_slice() == b"true")
    } else {
        Ok(true)
    }
}

#[inline]
fn parse_u32_attr(e: &BytesStart, attr_name: &[u8]) -> Option<u32> {
    get_attr(e, attr_name).and_then(|v| std::str::from_utf8(&v).ok()?.parse().ok())
}

#[inline]
fn get_attr(e: &BytesStart, name: &[u8]) -> Option<Vec<u8>> {
    e.attributes()
        .filter_map(|a| a.ok())
        .find(|a| a.key.as_ref() == name)
        .map(|a| a.value.to_vec())
}
