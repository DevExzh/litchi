//! Chart XML reader.
//!
//! This module provides functionality to parse chart XML files
//! from OOXML packages.

use crate::charts::axis::{Axis, CategoryAxis, DateAxis, SeriesAxis, ValueAxis};
use crate::charts::chart::{Chart, View3D, WallFloor};
use crate::charts::legend::Legend;
use crate::charts::models::{NumericData, RichText, StringData, TitleText};
use crate::charts::plot_area::{
    AreaTypeGroup, Bar3DTypeGroup, BarTypeGroup, LineTypeGroup, PieTypeGroup, PlotArea,
    ScatterTypeGroup, TypeGroup, TypeGroupCommon,
};
use crate::charts::series::Series;
use crate::charts::types::{
    AxisPosition, BarDirection, BarGrouping, DisplayBlanks, LegendPosition, ScatterStyle,
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
                    b"scatterChart" => {
                        if let Some(group) = parse_scatter_chart(reader)? {
                            plot_area.type_groups.push(TypeGroup::Scatter(group));
                        }
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
                    b"barDir" => {
                        if let Some(val) = get_attr(e, b"val") {
                            direction = if val.as_slice() == b"bar" {
                                BarDirection::Bar
                            } else {
                                BarDirection::Column
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
                    b"barDir" => {
                        if let Some(val) = get_attr(e, b"val") {
                            direction = if val.as_slice() == b"bar" {
                                BarDirection::Bar
                            } else {
                                BarDirection::Column
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
                                _ => ScatterStyle::LineMarker,
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
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pt" => {
                if let Some(val) = parse_point_value(reader)? {
                    data.values.push(val);
                }
            },
            Ok(Event::End(ref e))
                if matches!(e.local_name().as_ref(), b"val" | b"xVal" | b"yVal") =>
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
                    <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>East</c:v></c:pt>
                        </c:strCache></c:strRef></c:cat>
                    <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>42.5</c:v></c:pt>
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
            group.common.series[0].values.as_ref().unwrap().values,
            [42.5]
        );
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
}
