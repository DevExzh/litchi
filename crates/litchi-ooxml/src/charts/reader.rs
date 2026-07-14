//! Chart XML reader.
//!
//! This module provides functionality to parse chart XML files
//! from OOXML packages.

use crate::charts::axis::{
    Axis, AxisCommon, AxisCrossBetween, AxisCrossMode, AxisLabelAlign, BuiltInUnit, CategoryAxis,
    DateAxis, DisplayUnits, SeriesAxis, TimeUnit, ValueAxis,
};
use crate::charts::chart::{Chart, View3D, WallFloor};
use crate::charts::legend::{Legend, LegendEntry};
use crate::charts::models::{
    DataSourceRef, Layout, NumberFormat, NumericData, RichText, StringData, TitleText,
};
use crate::charts::plot_area::{
    Area3DTypeGroup, AreaTypeGroup, Bar3DTypeGroup, BarShape, BarTypeGroup, BubbleTypeGroup,
    DoughnutTypeGroup, Line3DTypeGroup, LineTypeGroup, Pie3DTypeGroup, PieTypeGroup, PlotArea,
    RadarTypeGroup, ScatterTypeGroup, StockTypeGroup, Surface3DTypeGroup, SurfaceTypeGroup,
    TypeGroup, TypeGroupCommon,
};
use crate::charts::series::{
    DataLabel, DataLabels, DataPoint, ErrorBar, ErrorBarDirection, ErrorBarType, ErrorBarValueType,
    Series, Trendline, TrendlineType,
};
use crate::charts::types::{
    AxisOrientation, AxisPosition, BarDirection, BarGrouping, DataLabelPosition, DisplayBlanks,
    LayoutMode, LayoutTarget, LegendPosition, MarkerStyle, RadarStyle, ScatterStyle,
    TickLabelPosition, TickMark,
};
use crate::common::xml::{decode_xml_reference, is_drawingml_chart_name, is_drawingml_name};
use crate::error::{OoxmlError, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::{Config, NsReader};
use std::io::BufRead;

const IGNORED_NAMESPACE_ELEMENT: &str = "ignoredNamespaceElement";

/// Namespace-aware streaming adapter for the chart model parser.
///
/// Core chart elements are exposed unchanged, DrawingML text elements are kept so
/// rich titles can be decoded, and all other namespaces are skipped as extension
/// content. Rewriting the remaining DrawingML container names prevents them from
/// being mistaken for same-local-name chart elements by the focused parsers below.
struct ChartXmlReader<R: BufRead> {
    inner: NsReader<R>,
    depth: usize,
    skipped_depth: usize,
    saw_root: bool,
    closed_root: bool,
}

impl<R: BufRead> ChartXmlReader<R> {
    fn from_reader(reader: R) -> Self {
        Self {
            inner: NsReader::from_reader(reader),
            depth: 0,
            skipped_depth: 0,
            saw_root: false,
            closed_root: false,
        }
    }

    fn config_mut(&mut self) -> &mut Config {
        self.inner.config_mut()
    }

    fn decoder(&self) -> Decoder {
        self.inner.decoder()
    }

    fn read_event_into<'buffer>(&mut self, buffer: &'buffer mut Vec<u8>) -> Result<Event<'buffer>> {
        let (namespace, event) = self
            .inner
            .read_resolved_event_into(buffer)
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;

        match event {
            Event::Start(mut element) => {
                let is_chart = is_chart_namespace(&namespace, &element);
                let is_drawing = is_drawing_namespace(&namespace, &element);

                if self.depth == 0 {
                    if self.saw_root
                        || !is_drawingml_chart_name(&namespace, element.name(), b"chartSpace")
                    {
                        return Err(OoxmlError::InvalidFormat(
                            "chart XML must have one DrawingML chartSpace root".to_string(),
                        ));
                    }
                    self.saw_root = true;
                }
                self.depth = self.depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("chart XML nesting is too deep".to_string())
                })?;

                if self.skipped_depth > 0 || (!is_chart && !is_drawing) {
                    self.skipped_depth = self.skipped_depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("chart XML nesting is too deep".to_string())
                    })?;
                    element.set_name(IGNORED_NAMESPACE_ELEMENT.as_bytes());
                } else if is_drawing && element.local_name().as_ref() != b"t" {
                    element.set_name(IGNORED_NAMESPACE_ELEMENT.as_bytes());
                }
                Ok(Event::Start(element))
            },
            Event::Empty(mut element) => {
                if self.depth == 0 {
                    return Err(OoxmlError::InvalidFormat(
                        "chart XML must have one non-empty DrawingML chartSpace root".to_string(),
                    ));
                }
                let is_chart = is_chart_namespace(&namespace, &element);
                let is_drawing = is_drawing_namespace(&namespace, &element);
                if self.skipped_depth > 0
                    || (!is_chart && (!is_drawing || element.local_name().as_ref() != b"t"))
                {
                    element.set_name(IGNORED_NAMESPACE_ELEMENT.as_bytes());
                }
                Ok(Event::Empty(element))
            },
            Event::End(element) => {
                if self.depth == 0 {
                    return Err(OoxmlError::InvalidFormat(
                        "chart XML has an unmatched closing element".to_string(),
                    ));
                }
                self.depth -= 1;
                if self.skipped_depth > 0 {
                    self.skipped_depth -= 1;
                    return Ok(Event::End(BytesEnd::new(IGNORED_NAMESPACE_ELEMENT)));
                }

                let is_chart = is_drawingml_chart_name(
                    &namespace,
                    element.name(),
                    element.local_name().as_ref(),
                );
                let is_drawing =
                    is_drawingml_name(&namespace, element.name(), element.local_name().as_ref());
                if self.depth == 0 {
                    if !is_drawingml_chart_name(&namespace, element.name(), b"chartSpace") {
                        return Err(OoxmlError::InvalidFormat(
                            "chart XML has an invalid root closing element".to_string(),
                        ));
                    }
                    self.closed_root = true;
                }
                if is_drawing && element.local_name().as_ref() != b"t" {
                    return Ok(Event::End(BytesEnd::new(IGNORED_NAMESPACE_ELEMENT)));
                }
                if is_chart || is_drawing {
                    return Ok(Event::End(element));
                }
                Err(OoxmlError::InvalidFormat(
                    "chart XML namespace state is inconsistent".to_string(),
                ))
            },
            _ if self.skipped_depth > 0 => Ok(Event::Comment(BytesText::new(""))),
            Event::Text(ref text) if self.depth == 0 => {
                if !text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    return Err(OoxmlError::InvalidFormat(
                        "chart XML contains text outside its root".to_string(),
                    ));
                }
                Ok(event)
            },
            Event::CData(_) | Event::GeneralRef(_) if self.depth == 0 => Err(
                OoxmlError::InvalidFormat("chart XML contains data outside its root".to_string()),
            ),
            Event::Eof if !self.saw_root || !self.closed_root => Err(OoxmlError::InvalidFormat(
                "chart XML has no complete chartSpace root".to_string(),
            )),
            Event::Eof => Ok(Event::Eof),
            _ => Ok(event),
        }
    }
}

fn is_chart_namespace(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    is_drawingml_chart_name(namespace, element.name(), element.local_name().as_ref())
}

fn is_drawing_namespace(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    is_drawingml_name(namespace, element.name(), element.local_name().as_ref())
}

/// Parse a chart XML document.
pub fn parse_chart<R: BufRead>(reader: R) -> Result<Chart> {
    let mut xml_reader = ChartXmlReader::from_reader(reader);
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
                        if chart.title.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "chart contains duplicate titles".into(),
                            ));
                        }
                        let title = parse_title(&mut xml_reader)?;
                        chart.title = Some(title.text);
                        chart.title_layout = title.layout;
                        chart.title_overlay = title.overlay;
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
                    b"showDLblsOverMax" => {
                        chart.show_data_labels_over_max = parse_bool_attr(e)?;
                    },
                    b"date1904" => {
                        chart.date_1904 = parse_bool_attr(e)?;
                    },
                    b"roundedCorners" => {
                        chart.rounded_corners = parse_bool_attr(e)?;
                    },
                    b"style" => {
                        chart.style = Some(bounded_u32_attr(e, "chart style", 1, 48)?);
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(chart)
}

struct ParsedTitle {
    text: TitleText,
    layout: Option<Layout>,
    overlay: bool,
}

fn parse_title<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<ParsedTitle> {
    let mut text = String::new();
    let mut formula = None;
    let mut layout = None;
    let mut overlay = false;
    let mut saw_overlay = false;
    let mut buf = Vec::new();
    let mut in_text = false;
    let mut saw_text = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if layout.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart title contains duplicate layouts".into(),
                    ));
                }
                layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                layout = Some(match layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(OoxmlError::InvalidFormat(
                            "chart title contains duplicate layouts".into(),
                        ));
                    },
                });
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"overlay" =>
            {
                if saw_overlay {
                    return Err(OoxmlError::InvalidFormat(
                        "chart title contains duplicate overlay flags".into(),
                    ));
                }
                overlay = parse_bool_attr(element)?;
                saw_overlay = true;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"f" => {
                if formula.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart title contains duplicate formula references".to_string(),
                    ));
                }
                formula = Some(parse_text_element(reader, b"f")?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"t" => {
                saw_text = true;
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let text = if let Some(formula) = formula {
        if saw_text {
            return Err(OoxmlError::InvalidFormat(
                "chart title mixes a formula reference with literal text".to_string(),
            ));
        }
        TitleText::Reference(DataSourceRef::new(formula))
    } else {
        TitleText::Literal(RichText::new(text))
    };
    Ok(ParsedTitle {
        text,
        layout,
        overlay,
    })
}

fn parse_view_3d<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<View3D> {
    let mut view = View3D::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"rotX" => {
                        view.rot_x = Some(bounded_u32_attr(e, "chart X rotation", 0, 90)?);
                    },
                    b"rotY" => {
                        view.rot_y = Some(bounded_u32_attr(e, "chart Y rotation", 0, 360)?);
                    },
                    b"perspective" => {
                        view.perspective = Some(bounded_u32_attr(e, "chart perspective", 0, 240)?);
                    },
                    b"hPercent" => {
                        view.height_percent =
                            Some(bounded_u32_attr(e, "chart height percentage", 5, 500)?);
                    },
                    b"depthPercent" => {
                        view.depth_percent =
                            Some(bounded_u32_attr(e, "chart depth percentage", 20, 2000)?);
                    },
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(view)
}

fn parse_wall_floor<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<WallFloor> {
    let mut wall_floor = WallFloor::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"thickness" =>
            {
                wall_floor.thickness = Some(bounded_u32_attr(e, "chart wall thickness", 0, 4096)?);
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(wall_floor)
}

fn parse_plot_area<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<PlotArea> {
    let mut plot_area = PlotArea::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"layout" => {
                plot_area.layout = Some(parse_layout(reader)?);
            },
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(plot_area)
}

fn parse_layout<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Layout> {
    let mut layout = Layout::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"layoutTarget" => {
                    layout.target = Some(
                        match required_enum_attr(e, "chart layout target")?.as_str() {
                            "inner" => LayoutTarget::Inner,
                            "outer" => LayoutTarget::Outer,
                            value => {
                                return Err(invalid_attribute(
                                    "chart layout target",
                                    value.as_bytes(),
                                ));
                            },
                        },
                    );
                },
                b"xMode" => layout.x_mode = Some(parse_layout_mode(e)?),
                b"yMode" => layout.y_mode = Some(parse_layout_mode(e)?),
                b"wMode" => layout.width_mode = Some(parse_layout_mode(e)?),
                b"hMode" => layout.height_mode = Some(parse_layout_mode(e)?),
                b"x" => layout.x = Some(required_f64_attr(e, "chart layout X position")?),
                b"y" => layout.y = Some(required_f64_attr(e, "chart layout Y position")?),
                b"w" => layout.width = Some(required_f64_attr(e, "chart layout width")?),
                b"h" => layout.height = Some(required_f64_attr(e, "chart layout height")?),
                _ => {},
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"layout" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart layout".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(layout)
}

fn parse_layout_mode(element: &BytesStart<'_>) -> Result<LayoutMode> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute("chart layout mode"))?;
    match value.as_slice() {
        b"edge" => Ok(LayoutMode::Edge),
        b"factor" => Ok(LayoutMode::Factor),
        _ => Err(invalid_attribute("chart layout mode", &value)),
    }
}

fn parse_common_type_group<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(common)
}

fn parse_area_3d_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Area3DTypeGroup> {
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

fn parse_bubble_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<BubbleTypeGroup> {
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

fn parse_doughnut_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DoughnutTypeGroup> {
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

fn parse_line_3d_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Line3DTypeGroup> {
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

fn parse_pie_3d_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Pie3DTypeGroup> {
    let mut group = Pie3DTypeGroup::new();
    group.common = parse_common_type_group(reader, b"pie3DChart", |_| Ok(()))?;
    Ok(group)
}

fn parse_radar_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<RadarTypeGroup> {
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

fn parse_stock_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<StockTypeGroup> {
    let mut group = StockTypeGroup::new();
    group.common = parse_common_type_group(reader, b"stockChart", |_| Ok(()))?;
    Ok(group)
}

fn parse_surface_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<SurfaceTypeGroup> {
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

fn parse_surface_3d_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Surface3DTypeGroup> {
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

fn parse_bar_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<BarTypeGroup>> {
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
            Err(error) => return Err(error),
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

fn parse_bar_3d_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<Bar3DTypeGroup>> {
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
            Err(error) => return Err(error),
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

fn parse_line_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<LineTypeGroup>> {
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = LineTypeGroup::new(grouping);
    group.common = common;
    group.marker = marker;
    Ok(Some(group))
}

fn parse_pie_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<PieTypeGroup>> {
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = PieTypeGroup::new();
    group.common = common;
    group.first_slice_angle = first_slice_angle;
    Ok(Some(group))
}

fn parse_area_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<AreaTypeGroup>> {
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = AreaTypeGroup::new(grouping);
    group.common = common;
    Ok(Some(group))
}

fn parse_scatter_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<ScatterTypeGroup>> {
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = ScatterTypeGroup::new(style);
    group.common = common;
    Ok(Some(group))
}

fn parse_series<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<Series>> {
    let mut series = Series::new(0);
    let mut saw_index = false;
    let mut saw_order = false;
    let mut saw_marker = false;
    let mut saw_data_labels = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"idx" => {
                        if saw_index {
                            return Err(OoxmlError::InvalidFormat(
                                "chart series has duplicate index".to_string(),
                            ));
                        }
                        saw_index = true;
                        series.index = required_u32_attr(e, "chart series index")?;
                    },
                    b"order" => {
                        if saw_order {
                            return Err(OoxmlError::InvalidFormat(
                                "chart series has duplicate order".to_string(),
                            ));
                        }
                        saw_order = true;
                        series.order = required_u32_attr(e, "chart series order")?;
                    },
                    b"tx" => {
                        series.title = parse_series_title(reader)?;
                    },
                    b"marker" => {
                        if saw_marker {
                            return Err(OoxmlError::InvalidFormat(
                                "chart series has duplicate marker".to_string(),
                            ));
                        }
                        saw_marker = true;
                        (series.marker_symbol, series.marker_size) = parse_series_marker(reader)?;
                    },
                    b"invertIfNegative" => {
                        series.invert_if_negative = parse_bool_attr(e)?;
                    },
                    b"dPt" => {
                        let point = parse_data_point(reader)?;
                        if series
                            .data_points
                            .iter()
                            .any(|existing| existing.index == point.index)
                        {
                            return Err(OoxmlError::InvalidFormat(format!(
                                "chart series has duplicate data-point index {}",
                                point.index
                            )));
                        }
                        series.data_points.push(point);
                    },
                    b"dLbls" => {
                        if saw_data_labels {
                            return Err(OoxmlError::InvalidFormat(
                                "chart series has duplicate data-label settings".to_string(),
                            ));
                        }
                        saw_data_labels = true;
                        series.data_labels = Some(parse_data_labels(reader)?);
                    },
                    b"trendline" => series.trendlines.push(parse_trendline(reader)?),
                    b"errBars" => {
                        let error_bar = parse_error_bar(reader)?;
                        if series
                            .error_bars
                            .iter()
                            .any(|existing| existing.direction == error_bar.direction)
                        {
                            return Err(OoxmlError::InvalidFormat(
                                "chart series has duplicate error-bar direction".to_string(),
                            ));
                        }
                        series.error_bars.push(error_bar);
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
                    b"smooth" => {
                        series.smooth = parse_bool_attr(e)?;
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if !saw_index || !saw_order {
        return Err(OoxmlError::InvalidFormat(
            "chart series requires both index and order".to_string(),
        ));
    }
    Ok(Some(series))
}

fn parse_series_marker<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<(Option<MarkerStyle>, Option<u32>)> {
    let mut symbol = None;
    let mut size = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"symbol" => {
                        if symbol.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "chart marker has duplicate symbol".to_string(),
                            ));
                        }
                        symbol = Some(parse_marker_style(element)?);
                    },
                    b"size" => {
                        if size.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "chart marker has duplicate size".to_string(),
                            ));
                        }
                        size = Some(bounded_u32_attr(element, "chart marker size", 2, 72)?);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"marker" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart series marker".to_string(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok((symbol, size))
}

fn parse_data_point<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DataPoint> {
    let mut index = None;
    let mut explosion = None;
    let mut marker_size = None;
    let mut marker_symbol = None;
    let mut invert_if_negative = false;
    let mut bubble_3d = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"marker" => {
                (marker_symbol, marker_size) = parse_series_marker(reader)?;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"idx" => {
                        index = Some(required_u32_attr(element, "chart data-point index")?);
                    },
                    b"explosion" => {
                        explosion = Some(required_u32_attr(element, "chart data-point explosion")?);
                    },
                    b"invertIfNegative" => invert_if_negative = parse_bool_attr(element)?,
                    b"bubble3D" => bubble_3d = Some(parse_bool_attr(element)?),
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dPt" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart data point".to_string(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    let mut point =
        DataPoint::new(index.ok_or_else(|| missing_attribute("chart data-point index"))?);
    point.explosion = explosion;
    point.marker_size = marker_size;
    point.marker_symbol = marker_symbol;
    point.invert_if_negative = invert_if_negative;
    point.bubble_3d = bubble_3d;
    Ok(point)
}

fn parse_data_labels<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DataLabels> {
    let mut labels = DataLabels::new();
    let mut saw_number_format = false;
    let mut saw_delete = false;
    let mut saw_shared_settings = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dLbl" => {
                let label = parse_data_label(reader)?;
                if labels
                    .labels
                    .iter()
                    .any(|existing| existing.index == label.index)
                {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "chart data labels contain duplicate point index {}",
                        label.index
                    )));
                }
                labels.labels.push(label);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dLbl" => {
                return Err(OoxmlError::InvalidFormat(
                    "chart point data label is missing its index".into(),
                ));
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"separator" => {
                saw_shared_settings = true;
                labels.separator = Some(parse_text_element(reader, b"separator")?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"delete" => {
                        if saw_delete {
                            return Err(OoxmlError::InvalidFormat(
                                "chart data labels contain duplicate delete flags".into(),
                            ));
                        }
                        labels.deleted = parse_bool_attr(element)?;
                        saw_delete = true;
                    },
                    b"numFmt" => {
                        if saw_number_format {
                            return Err(OoxmlError::InvalidFormat(
                                "chart data labels contain duplicate number formats".into(),
                            ));
                        }
                        labels.number_format = Some(parse_number_format(
                            element,
                            reader.decoder(),
                            "chart data-label",
                        )?);
                        saw_number_format = true;
                        saw_shared_settings = true;
                    },
                    b"dLblPos" => {
                        labels.position = Some(parse_data_label_position(element)?);
                        saw_shared_settings = true;
                    },
                    b"showLegendKey" => {
                        labels.show_legend_key = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    b"showVal" => {
                        labels.show_value = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    b"showCatName" => {
                        labels.show_category_name = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    b"showSerName" => {
                        labels.show_series_name = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    b"showPercent" => {
                        labels.show_percent = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    b"showBubbleSize" => {
                        labels.show_bubble_size = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    b"showLeaderLines" => {
                        labels.show_leader_lines = parse_bool_attr(element)?;
                        saw_shared_settings = true;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dLbls" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart data labels".to_string(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    if saw_delete && saw_shared_settings {
        return Err(OoxmlError::InvalidFormat(
            "chart data labels mix deletion with shared settings".into(),
        ));
    }
    Ok(labels)
}

fn parse_data_label<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DataLabel> {
    let mut label = DataLabel::new(0);
    let mut saw_index = false;
    let mut saw_delete = false;
    let mut saw_settings = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if label.layout.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart point data label contains duplicate layouts".into(),
                    ));
                }
                label.layout = Some(parse_layout(reader)?);
                saw_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                label.layout = Some(match label.layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(OoxmlError::InvalidFormat(
                            "chart point data label contains duplicate layouts".into(),
                        ));
                    },
                });
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"tx" => {
                if label.text.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart point data label contains duplicate text".into(),
                    ));
                }
                label.text = parse_label_text(reader)?;
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"separator" => {
                label.separator = Some(parse_text_element(reader, b"separator")?);
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                let is_setting = match element.local_name().as_ref() {
                    b"idx" => {
                        if saw_index {
                            return Err(OoxmlError::InvalidFormat(
                                "chart point data label contains duplicate indexes".into(),
                            ));
                        }
                        label.index = required_u32_attr(element, "chart point data-label index")?;
                        saw_index = true;
                        false
                    },
                    b"delete" => {
                        if saw_delete {
                            return Err(OoxmlError::InvalidFormat(
                                "chart point data label contains duplicate delete flags".into(),
                            ));
                        }
                        label.deleted = parse_bool_attr(element)?;
                        saw_delete = true;
                        false
                    },
                    b"numFmt" => {
                        if label.number_format.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "chart point data label contains duplicate number formats".into(),
                            ));
                        }
                        label.number_format = Some(parse_number_format(
                            element,
                            reader.decoder(),
                            "chart point data-label",
                        )?);
                        true
                    },
                    b"dLblPos" => {
                        label.position = Some(parse_data_label_position(element)?);
                        true
                    },
                    b"showLegendKey" => {
                        label.show_legend_key = parse_bool_attr(element)?;
                        true
                    },
                    b"showVal" => {
                        label.show_value = parse_bool_attr(element)?;
                        true
                    },
                    b"showCatName" => {
                        label.show_category_name = parse_bool_attr(element)?;
                        true
                    },
                    b"showSerName" => {
                        label.show_series_name = parse_bool_attr(element)?;
                        true
                    },
                    b"showPercent" => {
                        label.show_percent = parse_bool_attr(element)?;
                        true
                    },
                    b"showBubbleSize" => {
                        label.show_bubble_size = parse_bool_attr(element)?;
                        true
                    },
                    _ => false,
                };
                saw_settings |= is_setting;
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dLbl" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart point data label".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if !saw_index {
        return Err(missing_attribute("chart point data-label index"));
    }
    if saw_delete && saw_settings {
        return Err(OoxmlError::InvalidFormat(
            "chart point data label mixes deletion with label settings".into(),
        ));
    }
    Ok(label)
}

fn parse_label_text<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<TitleText>> {
    let mut text = String::new();
    let mut formula = None;
    let mut saw_text = false;
    let mut in_text = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"f" => {
                if formula.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data label contains duplicate formula references".into(),
                    ));
                }
                formula = Some(parse_text_element(reader, b"f")?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"t" => {
                saw_text = true;
                in_text = true;
            },
            Ok(Event::Text(value)) if in_text => text.push_str(
                &value
                    .decode()
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?,
            ),
            Ok(Event::CData(value)) if in_text => text.push_str(
                &value
                    .decode()
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?,
            ),
            Ok(Event::GeneralRef(reference)) if in_text => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"t" => {
                in_text = false;
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"tx" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart data-label text".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if let Some(formula) = formula {
        if saw_text {
            return Err(OoxmlError::InvalidFormat(
                "chart data label mixes a formula reference with literal text".into(),
            ));
        }
        Ok(Some(TitleText::Reference(DataSourceRef::new(formula))))
    } else if saw_text {
        Ok(Some(TitleText::Literal(RichText::new(text))))
    } else {
        Ok(None)
    }
}

fn parse_trendline<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Trendline> {
    let mut trendline = Trendline::linear();
    let mut saw_type = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"trendlineName" => {
                trendline.name = Some(parse_text_element(reader, b"trendlineName")?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"trendlineLbl" => {
                if trendline.show_label {
                    return Err(OoxmlError::InvalidFormat(
                        "chart trendline contains duplicate labels".into(),
                    ));
                }
                trendline.show_label = true;
                (
                    trendline.label,
                    trendline.label_layout,
                    trendline.label_number_format,
                ) = parse_trendline_label(reader)?;
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"trendlineLbl" => {
                if trendline.show_label {
                    return Err(OoxmlError::InvalidFormat(
                        "chart trendline contains duplicate labels".into(),
                    ));
                }
                trendline.show_label = true;
            },
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"trendlineType" => {
                    let value = get_attr(e, b"val")
                        .ok_or_else(|| missing_attribute("chart trendline type"))?;
                    trendline.trendline_type = match value.as_slice() {
                        b"exp" => TrendlineType::Exponential,
                        b"linear" => TrendlineType::Linear,
                        b"log" => TrendlineType::Logarithmic,
                        b"movingAvg" => TrendlineType::MovingAverage,
                        b"poly" => TrendlineType::Polynomial,
                        b"power" => TrendlineType::Power,
                        _ => return Err(invalid_attribute("chart trendline type", &value)),
                    };
                    saw_type = true;
                },
                b"order" => trendline.order = Some(bounded_u32_attr(e, "trendline order", 2, 6)?),
                b"period" => {
                    trendline.period = Some(bounded_u32_attr(e, "trendline period", 2, 255)?)
                },
                b"forward" => {
                    trendline.forward = Some(required_nonnegative_f64_attr(e, "trendline forward")?)
                },
                b"backward" => {
                    trendline.backward =
                        Some(required_nonnegative_f64_attr(e, "trendline backward")?)
                },
                b"intercept" => {
                    trendline.intercept = Some(required_f64_attr(e, "trendline intercept")?)
                },
                b"dispEq" => trendline.display_equation = parse_bool_attr(e)?,
                b"dispRSqr" => trendline.display_r_squared = parse_bool_attr(e)?,
                _ => {},
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"trendline" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart trendline".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    if !saw_type {
        return Err(missing_attribute("chart trendline type"));
    }
    if matches!(trendline.trendline_type, TrendlineType::Polynomial) && trendline.order.is_none() {
        return Err(missing_attribute("polynomial trendline order"));
    }
    if matches!(trendline.trendline_type, TrendlineType::MovingAverage)
        && trendline.period.is_none()
    {
        return Err(missing_attribute("moving-average trendline period"));
    }
    if !matches!(trendline.trendline_type, TrendlineType::Polynomial) && trendline.order.is_some() {
        return Err(OoxmlError::InvalidFormat(
            "only polynomial trendlines can specify an order".to_string(),
        ));
    }
    if !matches!(trendline.trendline_type, TrendlineType::MovingAverage)
        && trendline.period.is_some()
    {
        return Err(OoxmlError::InvalidFormat(
            "only moving-average trendlines can specify a period".to_string(),
        ));
    }
    Ok(trendline)
}

fn parse_trendline_label<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<(Option<TitleText>, Option<Layout>, Option<NumberFormat>)> {
    let mut text = None;
    let mut saw_text = false;
    let mut layout = None;
    let mut number_format = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if layout.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart trendline label contains duplicate layouts".into(),
                    ));
                }
                layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                layout = Some(match layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(OoxmlError::InvalidFormat(
                            "chart trendline label contains duplicate layouts".into(),
                        ));
                    },
                });
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"tx" => {
                if saw_text {
                    return Err(OoxmlError::InvalidFormat(
                        "chart trendline label contains duplicate text".into(),
                    ));
                }
                saw_text = true;
                text = parse_label_text(reader)?;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"numFmt" =>
            {
                if number_format.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart trendline label contains duplicate number formats".into(),
                    ));
                }
                number_format = Some(parse_number_format(
                    element,
                    reader.decoder(),
                    "chart trendline-label",
                )?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"trendlineLbl" => {
                break;
            },
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart trendline label".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok((text, layout, number_format))
}

fn parse_error_bar<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<ErrorBar> {
    let mut direction = None;
    let mut error_type = None;
    let mut value_type = None;
    let mut value = None;
    let mut plus_values = None;
    let mut minus_values = None;
    let mut no_end_cap = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"plus" => {
                plus_values = parse_numeric_data(reader)?
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"minus" => {
                minus_values = parse_numeric_data(reader)?
            },
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"errDir" => {
                    direction = Some(
                        match required_enum_attr(e, "error-bar direction")?.as_str() {
                            "x" => ErrorBarDirection::X,
                            "y" => ErrorBarDirection::Y,
                            value => {
                                return Err(invalid_attribute(
                                    "error-bar direction",
                                    value.as_bytes(),
                                ));
                            },
                        },
                    )
                },
                b"errBarType" => {
                    error_type = Some(match required_enum_attr(e, "error-bar type")?.as_str() {
                        "both" => ErrorBarType::Both,
                        "plus" => ErrorBarType::Plus,
                        "minus" => ErrorBarType::Minus,
                        value => return Err(invalid_attribute("error-bar type", value.as_bytes())),
                    })
                },
                b"errValType" => {
                    value_type = Some(
                        match required_enum_attr(e, "error-bar value type")?.as_str() {
                            "fixedVal" => ErrorBarValueType::Fixed,
                            "percentage" => ErrorBarValueType::Percentage,
                            "stdDev" => ErrorBarValueType::StdDev,
                            "stdErr" => ErrorBarValueType::StdErr,
                            "cust" => ErrorBarValueType::Custom,
                            value => {
                                return Err(invalid_attribute(
                                    "error-bar value type",
                                    value.as_bytes(),
                                ));
                            },
                        },
                    )
                },
                b"noEndCap" => no_end_cap = parse_bool_attr(e)?,
                b"val" => value = Some(required_nonnegative_f64_attr(e, "error-bar value")?),
                _ => {},
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"errBars" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart error bars".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    let error_bar = ErrorBar {
        direction: direction.ok_or_else(|| missing_attribute("error-bar direction"))?,
        error_type: error_type.ok_or_else(|| missing_attribute("error-bar type"))?,
        value_type: value_type.ok_or_else(|| missing_attribute("error-bar value type"))?,
        value,
        plus_values,
        minus_values,
        no_end_cap,
    };
    match error_bar.value_type {
        ErrorBarValueType::Fixed | ErrorBarValueType::Percentage | ErrorBarValueType::StdDev
            if error_bar.value.is_none() =>
        {
            return Err(missing_attribute("error-bar scalar value"));
        },
        ErrorBarValueType::Custom
            if error_bar.plus_values.is_none() && error_bar.minus_values.is_none() =>
        {
            return Err(missing_attribute("custom error-bar values"));
        },
        ErrorBarValueType::StdErr | ErrorBarValueType::Custom if error_bar.value.is_some() => {
            return Err(OoxmlError::InvalidFormat(
                "standard-error and custom error bars cannot have a scalar value".to_string(),
            ));
        },
        _ => {},
    }
    Ok(error_bar)
}

fn parse_string_data<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<StringData>> {
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(data))
}

fn parse_numeric_data<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<NumericData>> {
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
                    b"val" | b"xVal" | b"yVal" | b"bubbleSize" | b"plus" | b"minus"
                ) =>
            {
                break;
            },
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart element".to_string(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(data))
}

fn parse_series_title<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<TitleText>> {
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
            Err(error) => return Err(error),
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

fn parse_text_element<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    end_name: &[u8],
) -> Result<String> {
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
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == IGNORED_NAMESPACE_ELEMENT.as_bytes() => {},
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(text)
}

fn parse_point_text<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<String>> {
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(Some(text))
}

fn parse_point_value<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<f64>> {
    if let Some(text) = parse_point_text(reader)? {
        Ok(Some(text.trim().parse::<f64>().map_err(|_| {
            OoxmlError::InvalidFormat(format!("invalid chart numeric point '{text}'"))
        })?))
    } else {
        Ok(None)
    }
}

struct ParsedAxisCommon {
    axis_id: Option<u32>,
    cross_axis_id: Option<u32>,
    position: Option<AxisPosition>,
    title: Option<TitleText>,
    title_layout: Option<Layout>,
    title_overlay: bool,
    number_format: Option<NumberFormat>,
    orientation: AxisOrientation,
    major_tick_mark: TickMark,
    minor_tick_mark: TickMark,
    tick_label_position: TickLabelPosition,
    deleted: bool,
    cross_mode: AxisCrossMode,
    crosses_at: Option<f64>,
    show_major_gridlines: bool,
    show_minor_gridlines: bool,
}

impl ParsedAxisCommon {
    fn new() -> Self {
        Self {
            axis_id: None,
            cross_axis_id: None,
            position: None,
            title: None,
            title_layout: None,
            title_overlay: false,
            number_format: None,
            orientation: AxisOrientation::MinMax,
            major_tick_mark: TickMark::Out,
            minor_tick_mark: TickMark::None,
            tick_label_position: TickLabelPosition::NextTo,
            deleted: false,
            cross_mode: AxisCrossMode::AutoZero,
            crosses_at: None,
            show_major_gridlines: false,
            show_minor_gridlines: false,
        }
    }

    fn finish(self) -> Result<AxisCommon> {
        let axis_id = self
            .axis_id
            .ok_or_else(|| missing_attribute("chart axis ID"))?;
        let position = self
            .position
            .ok_or_else(|| missing_attribute("chart axis position"))?;
        let cross_axis_id = self
            .cross_axis_id
            .ok_or_else(|| missing_attribute("chart crossing-axis ID"))?;
        let mut common = AxisCommon::new(axis_id, position, cross_axis_id);
        common.title = self.title;
        common.layout = self.title_layout;
        common.title_overlay = self.title_overlay;
        common.number_format = self.number_format;
        common.orientation = self.orientation;
        common.major_tick_mark = self.major_tick_mark;
        common.minor_tick_mark = self.minor_tick_mark;
        common.tick_label_position = self.tick_label_position;
        common.deleted = self.deleted;
        common.cross_mode = self.cross_mode;
        common.crosses_at = self.crosses_at;
        common.show_major_gridlines = self.show_major_gridlines;
        common.show_minor_gridlines = self.show_minor_gridlines;
        Ok(common)
    }
}

fn parse_axis_title<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    common: &mut ParsedAxisCommon,
) -> Result<()> {
    if common.title.is_some() {
        return Err(OoxmlError::InvalidFormat(
            "chart axis contains duplicate titles".into(),
        ));
    }
    let title = parse_title(reader)?;
    common.title = Some(title.text);
    common.title_layout = title.layout;
    common.title_overlay = title.overlay;
    Ok(())
}

fn parse_axis_common_element(
    common: &mut ParsedAxisCommon,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<bool> {
    match element.local_name().as_ref() {
        b"axId" => common.axis_id = Some(required_u32_attr(element, "chart axis ID")?),
        b"orientation" => common.orientation = parse_axis_orientation(element)?,
        b"delete" => common.deleted = parse_bool_attr(element)?,
        b"axPos" => common.position = Some(parse_axis_position(element)?),
        b"majorGridlines" => common.show_major_gridlines = true,
        b"minorGridlines" => common.show_minor_gridlines = true,
        b"numFmt" => {
            common.number_format = Some(parse_number_format(element, decoder, "chart axis")?)
        },
        b"majorTickMark" => common.major_tick_mark = parse_tick_mark(element)?,
        b"minorTickMark" => common.minor_tick_mark = parse_tick_mark(element)?,
        b"tickLblPos" => common.tick_label_position = parse_tick_label_position(element)?,
        b"crossAx" => {
            common.cross_axis_id = Some(required_u32_attr(element, "chart crossing-axis ID")?);
        },
        b"crosses" => common.cross_mode = parse_axis_cross_mode(element)?,
        b"crossesAt" => {
            common.crosses_at = Some(required_f64_attr(element, "chart axis crossing value")?);
        },
        _ => return Ok(false),
    }
    Ok(true)
}

fn parse_category_axis<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<CategoryAxis>> {
    let mut common = ParsedAxisCommon::new();
    let mut auto = true;
    let mut label_align = None;
    let mut label_offset = None;
    let mut tick_label_skip = None;
    let mut tick_mark_skip = None;
    let mut no_multi_level = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"title" => {
                parse_axis_title(reader, &mut common)?;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if !parse_axis_common_element(&mut common, element, reader.decoder())? =>
            {
                match element.local_name().as_ref() {
                    b"auto" => auto = parse_bool_attr(element)?,
                    b"lblAlgn" => label_align = Some(parse_axis_label_align(element)?),
                    b"lblOffset" => {
                        label_offset = Some(bounded_u32_attr(
                            element,
                            "chart axis label offset",
                            0,
                            1000,
                        )?);
                    },
                    b"tickLblSkip" => {
                        tick_label_skip = Some(required_positive_u32_attr(
                            element,
                            "chart tick-label skip",
                        )?);
                    },
                    b"tickMarkSkip" => {
                        tick_mark_skip =
                            Some(required_positive_u32_attr(element, "chart tick-mark skip")?);
                    },
                    b"noMultiLvlLbl" => no_multi_level = parse_bool_attr(element)?,
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"catAx" => break,
            Ok(Event::Eof) => return Err(unterminated_axis("category")),
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let common = common.finish()?;
    let mut axis = CategoryAxis::new(common.axis_id, common.position, common.cross_axis_id);
    axis.common = common;
    axis.auto = auto;
    axis.label_align = label_align;
    axis.label_offset = label_offset;
    axis.tick_label_skip = tick_label_skip;
    axis.tick_mark_skip = tick_mark_skip;
    axis.no_multi_level = no_multi_level;
    Ok(Some(axis))
}

fn parse_value_axis<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<ValueAxis>> {
    let mut common = ParsedAxisCommon::new();
    let mut min = None;
    let mut max = None;
    let mut major_unit = None;
    let mut minor_unit = None;
    let mut log_base = None;
    let mut display_units = None;
    let mut cross_between = AxisCrossBetween::Between;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"title" => {
                parse_axis_title(reader, &mut common)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dispUnits" => {
                if display_units.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart value axis contains duplicate display units".into(),
                    ));
                }
                display_units = Some(parse_display_units(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dispUnits" => {
                return Err(OoxmlError::InvalidFormat(
                    "chart display units are missing their unit".into(),
                ));
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if !parse_axis_common_element(&mut common, element, reader.decoder())? =>
            {
                match element.local_name().as_ref() {
                    b"min" => min = Some(required_f64_attr(element, "chart axis minimum")?),
                    b"max" => max = Some(required_f64_attr(element, "chart axis maximum")?),
                    b"majorUnit" => {
                        major_unit = Some(required_positive_f64_attr(
                            element,
                            "chart axis major unit",
                        )?);
                    },
                    b"minorUnit" => {
                        minor_unit = Some(required_positive_f64_attr(
                            element,
                            "chart axis minor unit",
                        )?);
                    },
                    b"logBase" => {
                        let value = required_f64_attr(element, "chart logarithmic base")?;
                        if !(2.0..=1000.0).contains(&value) {
                            return Err(OoxmlError::InvalidFormat(
                                "chart logarithmic base must be between 2 and 1000".into(),
                            ));
                        }
                        log_base = Some(value);
                    },
                    b"crossBetween" => cross_between = parse_axis_cross_between(element)?,
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"valAx" => break,
            Ok(Event::Eof) => return Err(unterminated_axis("value")),
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if min.zip(max).is_some_and(|(min, max)| min > max) {
        return Err(OoxmlError::InvalidFormat(
            "chart axis minimum exceeds maximum".into(),
        ));
    }
    let common = common.finish()?;
    let mut axis = ValueAxis::new(common.axis_id, common.position, common.cross_axis_id);
    axis.common = common;
    axis.min = min;
    axis.max = max;
    axis.major_unit = major_unit;
    axis.minor_unit = minor_unit;
    axis.log_base = log_base;
    axis.cross_between = cross_between;
    axis.display_units = display_units;
    Ok(Some(axis))
}

fn parse_display_units<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DisplayUnits> {
    let mut built_in_unit = None;
    let mut custom_unit = None;
    let mut label = None;
    let mut layout = None;
    let mut saw_label = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dispUnitsLbl" => {
                if saw_label {
                    return Err(OoxmlError::InvalidFormat(
                        "chart display units contain duplicate labels".into(),
                    ));
                }
                saw_label = true;
                (label, layout) = parse_display_units_label(reader)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dispUnitsLbl" => {
                if saw_label {
                    return Err(OoxmlError::InvalidFormat(
                        "chart display units contain duplicate labels".into(),
                    ));
                }
                saw_label = true;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"builtInUnit" => {
                        if built_in_unit.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "chart display units contain duplicate built-in units".into(),
                            ));
                        }
                        built_in_unit = Some(parse_built_in_unit(element)?);
                    },
                    b"custUnit" => {
                        if custom_unit.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "chart display units contain duplicate custom units".into(),
                            ));
                        }
                        custom_unit = Some(required_positive_f64_attr(
                            element,
                            "chart custom display unit",
                        )?);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dispUnits" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart display units".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if built_in_unit.is_some() == custom_unit.is_some() {
        return Err(OoxmlError::InvalidFormat(
            "chart display units require exactly one built-in or custom unit".into(),
        ));
    }
    Ok(DisplayUnits {
        built_in_unit,
        custom_unit,
        show_label: saw_label,
        label,
        layout,
    })
}

fn parse_display_units_label<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<(Option<TitleText>, Option<Layout>)> {
    let mut text = String::new();
    let mut formula = None;
    let mut layout = None;
    let mut in_text = false;
    let mut saw_text = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if layout.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart display-units label contains duplicate layouts".into(),
                    ));
                }
                layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                layout = Some(match layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(OoxmlError::InvalidFormat(
                            "chart display-units label contains duplicate layouts".into(),
                        ));
                    },
                });
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"f" => {
                if formula.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart display-units label contains duplicate formula references".into(),
                    ));
                }
                formula = Some(parse_text_element(reader, b"f")?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"t" => {
                saw_text = true;
                in_text = true;
            },
            Ok(Event::Text(value)) if in_text => {
                text.push_str(
                    &value
                        .decode()
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?,
                );
            },
            Ok(Event::CData(value)) if in_text => {
                text.push_str(
                    &value
                        .decode()
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?,
                );
            },
            Ok(Event::GeneralRef(reference)) if in_text => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"t" => {
                in_text = false;
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dispUnitsLbl" => {
                break;
            },
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart display-units label".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let label = if let Some(formula) = formula {
        if saw_text {
            return Err(OoxmlError::InvalidFormat(
                "chart display-units label mixes a formula reference with literal text".into(),
            ));
        }
        Some(TitleText::Reference(DataSourceRef::new(formula)))
    } else if saw_text {
        Some(TitleText::Literal(RichText::new(text)))
    } else {
        None
    };
    Ok((label, layout))
}

fn parse_date_axis<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<DateAxis>> {
    let mut common = ParsedAxisCommon::new();
    let mut min = None;
    let mut max = None;
    let mut major_unit = None;
    let mut minor_unit = None;
    let mut major_time_unit = None;
    let mut minor_time_unit = None;
    let mut base_time_unit = None;
    let mut auto = true;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"title" => {
                parse_axis_title(reader, &mut common)?;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if !parse_axis_common_element(&mut common, element, reader.decoder())? =>
            {
                match element.local_name().as_ref() {
                    b"min" => min = Some(required_f64_attr(element, "chart date minimum")?),
                    b"max" => max = Some(required_f64_attr(element, "chart date maximum")?),
                    b"majorUnit" => {
                        major_unit = Some(required_positive_f64_attr(
                            element,
                            "chart date-axis major unit",
                        )?);
                    },
                    b"minorUnit" => {
                        minor_unit = Some(required_positive_f64_attr(
                            element,
                            "chart date-axis minor unit",
                        )?);
                    },
                    b"majorTimeUnit" => major_time_unit = Some(parse_time_unit(element)?),
                    b"minorTimeUnit" => minor_time_unit = Some(parse_time_unit(element)?),
                    b"baseTimeUnit" => base_time_unit = Some(parse_time_unit(element)?),
                    b"auto" => auto = parse_bool_attr(element)?,
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dateAx" => break,
            Ok(Event::Eof) => return Err(unterminated_axis("date")),
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if min.zip(max).is_some_and(|(min, max)| min > max) {
        return Err(OoxmlError::InvalidFormat(
            "chart date-axis minimum exceeds maximum".into(),
        ));
    }
    let common = common.finish()?;
    let mut axis = DateAxis::new(common.axis_id, common.position, common.cross_axis_id);
    axis.common = common;
    axis.min = min;
    axis.max = max;
    axis.major_unit = major_unit;
    axis.minor_unit = minor_unit;
    axis.major_time_unit = major_time_unit;
    axis.minor_time_unit = minor_time_unit;
    axis.base_time_unit = base_time_unit;
    axis.auto = auto;
    Ok(Some(axis))
}

fn parse_series_axis<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<SeriesAxis>> {
    let mut common = ParsedAxisCommon::new();
    let mut tick_label_skip = None;
    let mut tick_mark_skip = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"title" => {
                parse_axis_title(reader, &mut common)?;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if !parse_axis_common_element(&mut common, element, reader.decoder())? =>
            {
                match element.local_name().as_ref() {
                    b"tickLblSkip" => {
                        tick_label_skip = Some(required_positive_u32_attr(
                            element,
                            "chart series-axis tick-label skip",
                        )?);
                    },
                    b"tickMarkSkip" => {
                        tick_mark_skip = Some(required_positive_u32_attr(
                            element,
                            "chart series-axis tick-mark skip",
                        )?);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"serAx" => break,
            Ok(Event::Eof) => return Err(unterminated_axis("series")),
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let common = common.finish()?;
    let mut axis = SeriesAxis::new(common.axis_id, common.position, common.cross_axis_id);
    axis.common = common;
    axis.tick_label_skip = tick_label_skip;
    axis.tick_mark_skip = tick_mark_skip;
    Ok(Some(axis))
}

fn unterminated_axis(kind: &str) -> OoxmlError {
    OoxmlError::InvalidFormat(format!("unterminated chart {kind} axis"))
}

fn parse_legend<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Legend> {
    let mut position = LegendPosition::Right;
    let mut overlay = false;
    let mut layout = None;
    let mut entries = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"legendEntry" => {
                let entry = parse_legend_entry(reader)?;
                if entries
                    .iter()
                    .any(|existing: &LegendEntry| existing.index == entry.index)
                {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "chart legend contains duplicate entry index {}",
                        entry.index
                    )));
                }
                entries.push(entry);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"legendEntry" => {
                return Err(OoxmlError::InvalidFormat(
                    "chart legend entry is missing its index".into(),
                ));
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if layout.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart legend contains duplicate layouts".into(),
                    ));
                }
                layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                layout = Some(match layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(OoxmlError::InvalidFormat(
                            "chart legend contains duplicate layouts".into(),
                        ));
                    },
                });
            },
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"legendPos" => {
                        let value = get_attr(e, b"val")
                            .ok_or_else(|| missing_attribute("chart legend position"))?;
                        position = match value.as_slice() {
                            b"b" => LegendPosition::Bottom,
                            b"l" => LegendPosition::Left,
                            b"r" => LegendPosition::Right,
                            b"t" => LegendPosition::Top,
                            b"tr" => LegendPosition::TopRight,
                            _ => return Err(invalid_attribute("chart legend position", &value)),
                        };
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
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut legend = Legend::new(position).with_overlay(overlay);
    legend.layout = layout;
    legend.entries = entries;
    Ok(legend)
}

fn parse_legend_entry<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<LegendEntry> {
    let mut index = None;
    let mut deleted = false;
    let mut saw_delete = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"idx" => {
                        if index.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "chart legend entry contains duplicate indexes".into(),
                            ));
                        }
                        index = Some(required_u32_attr(element, "chart legend entry index")?);
                    },
                    b"delete" => {
                        if saw_delete {
                            return Err(OoxmlError::InvalidFormat(
                                "chart legend entry contains duplicate delete flags".into(),
                            ));
                        }
                        deleted = parse_bool_attr(element)?;
                        saw_delete = true;
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"legendEntry" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart legend entry".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let index = index.ok_or_else(|| missing_attribute("chart legend entry index"))?;
    let mut entry = LegendEntry::new(index);
    entry.deleted = deleted;
    Ok(entry)
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

fn parse_axis_orientation(element: &BytesStart<'_>) -> Result<AxisOrientation> {
    let value =
        get_attr(element, b"val").ok_or_else(|| missing_attribute("chart axis orientation"))?;
    match value.as_slice() {
        b"minMax" => Ok(AxisOrientation::MinMax),
        b"maxMin" => Ok(AxisOrientation::MaxMin),
        _ => Err(invalid_attribute("chart axis orientation", &value)),
    }
}

fn parse_tick_mark(element: &BytesStart<'_>) -> Result<TickMark> {
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

fn parse_tick_label_position(element: &BytesStart<'_>) -> Result<TickLabelPosition> {
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

fn parse_axis_cross_mode(element: &BytesStart<'_>) -> Result<AxisCrossMode> {
    let value =
        get_attr(element, b"val").ok_or_else(|| missing_attribute("chart axis crossing mode"))?;
    match value.as_slice() {
        b"autoZero" => Ok(AxisCrossMode::AutoZero),
        b"max" => Ok(AxisCrossMode::Max),
        b"min" => Ok(AxisCrossMode::Min),
        _ => Err(invalid_attribute("chart axis crossing mode", &value)),
    }
}

fn parse_axis_cross_between(element: &BytesStart<'_>) -> Result<AxisCrossBetween> {
    let value = get_attr(element, b"val")
        .ok_or_else(|| missing_attribute("chart axis crossing position"))?;
    match value.as_slice() {
        b"between" => Ok(AxisCrossBetween::Between),
        b"midCat" => Ok(AxisCrossBetween::MidCategory),
        _ => Err(invalid_attribute("chart axis crossing position", &value)),
    }
}

fn parse_axis_label_align(element: &BytesStart<'_>) -> Result<AxisLabelAlign> {
    let value =
        get_attr(element, b"val").ok_or_else(|| missing_attribute("chart axis label alignment"))?;
    match value.as_slice() {
        b"ctr" => Ok(AxisLabelAlign::Center),
        b"l" => Ok(AxisLabelAlign::Left),
        b"r" => Ok(AxisLabelAlign::Right),
        _ => Err(invalid_attribute("chart axis label alignment", &value)),
    }
}

fn parse_time_unit(element: &BytesStart<'_>) -> Result<TimeUnit> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute("chart time unit"))?;
    match value.as_slice() {
        b"days" => Ok(TimeUnit::Days),
        b"months" => Ok(TimeUnit::Months),
        b"years" => Ok(TimeUnit::Years),
        _ => Err(invalid_attribute("chart time unit", &value)),
    }
}

fn parse_built_in_unit(element: &BytesStart<'_>) -> Result<BuiltInUnit> {
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

fn parse_marker_style(element: &BytesStart<'_>) -> Result<MarkerStyle> {
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

fn parse_data_label_position(element: &BytesStart<'_>) -> Result<DataLabelPosition> {
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

fn parse_number_format(
    element: &BytesStart<'_>,
    decoder: Decoder,
    description: &str,
) -> Result<NumberFormat> {
    let format_code = element
        .try_get_attribute(b"formatCode")
        .map_err(|error| OoxmlError::Xml(error.to_string()))?
        .ok_or_else(|| missing_attribute(&format!("{description} number format code")))?
        .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
        .map_err(|error| OoxmlError::Xml(error.to_string()))?
        .into_owned();
    let source_linked = match get_attr(element, b"sourceLinked") {
        Some(value) => parse_bool_value(&value, &format!("{description} source-linked flag"))?,
        None => true,
    };
    Ok(NumberFormat::new(format_code).with_source_linked(source_linked))
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
        parse_bool_value(&val, "chart boolean")
    } else {
        Ok(true)
    }
}

fn parse_bool_value(value: &[u8], description: &str) -> Result<bool> {
    match value {
        b"1" | b"true" => Ok(true),
        b"0" | b"false" => Ok(false),
        _ => Err(invalid_attribute(description, value)),
    }
}

fn invalid_attribute(description: &str, value: &[u8]) -> OoxmlError {
    OoxmlError::InvalidFormat(format!(
        "invalid {description} '{}'",
        String::from_utf8_lossy(value)
    ))
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

fn required_positive_u32_attr(element: &BytesStart<'_>, description: &str) -> Result<u32> {
    let value = required_u32_attr(element, description)?;
    if value == 0 {
        return Err(OoxmlError::InvalidFormat(format!(
            "{description} must be positive"
        )));
    }
    Ok(value)
}

fn required_f64_attr(element: &BytesStart<'_>, description: &str) -> Result<f64> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute(description))?;
    let parsed = std::str::from_utf8(&value)
        .ok()
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite())
        .ok_or_else(|| invalid_attribute(description, &value))?;
    Ok(parsed)
}

fn required_positive_f64_attr(element: &BytesStart<'_>, description: &str) -> Result<f64> {
    let value = required_f64_attr(element, description)?;
    if value <= 0.0 {
        return Err(OoxmlError::InvalidFormat(format!(
            "{description} must be positive"
        )));
    }
    Ok(value)
}

fn required_nonnegative_f64_attr(element: &BytesStart<'_>, description: &str) -> Result<f64> {
    let value = required_f64_attr(element, description)?;
    if value < 0.0 {
        return Err(OoxmlError::InvalidFormat(format!(
            "{description} must be nonnegative"
        )));
    }
    Ok(value)
}

fn required_enum_attr(element: &BytesStart<'_>, description: &str) -> Result<String> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute(description))?;
    String::from_utf8(value).map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
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
                    <c:dLbls><c:dLbl><c:idx val="0"/><c:delete val="1"/></c:dLbl>
                        <c:showVal val="1"/></c:dLbls>
                </c:ser></c:barChart></c:plotArea>
                <c:legend><c:legendPos val="b"/><c:overlay val="1"/></c:legend>
                <c:showDLblsOverMax val="1"/>
            </c:chart><c:style val="12"/></c:chartSpace>"#;

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
        let labels = group.common.series[0].data_labels.as_ref().unwrap();
        assert!(labels.show_value);
        assert!(!labels.deleted);
        assert_eq!(labels.labels.len(), 1);
        assert_eq!(labels.labels[0].index, 0);
        assert!(labels.labels[0].deleted);
        assert_eq!(chart.legend.unwrap().position, LegendPosition::Bottom);
        assert!(chart.show_data_labels_over_max);
        assert_eq!(chart.style, Some(12));
    }

    #[test]
    fn parses_strict_chart_and_drawingml_namespaces() {
        let xml = br#"<c:chartSpace xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart"
                xmlns:a="http://purl.oclc.org/ooxml/drawingml/main">
            <c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Strict title</a:t></a:r></a:p>
                </c:rich></c:tx></c:title><c:plotArea><c:pieChart></c:pieChart></c:plotArea></c:chart>
            <c:style val="7"/>
        </c:chartSpace>"#;

        let chart = parse_chart(xml.as_slice()).unwrap();
        let Some(TitleText::Literal(title)) = chart.title else {
            panic!("expected a literal chart title");
        };
        assert_eq!(title.text, "Strict title");
        assert_eq!(chart.style, Some(7));
        assert!(matches!(
            chart.plot_area.type_groups.as_slice(),
            [TypeGroup::Pie(_)]
        ));
    }

    #[test]
    fn ignores_foreign_namespace_lookalikes_and_their_descendants() {
        let xml = br#"<c:chartSpace
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:x="urn:example:chart-extension">
            <x:style val="48"/>
            <c:chart><c:title><c:tx><c:strRef>
                <c:f>Sheet<x:payload><c:style val="47"/>ignored</x:payload>1!$A$1</c:f>
            </c:strRef></c:tx></c:title><c:plotArea>
                <x:barChart><c:barChart><c:ser><c:idx val="9"/></c:ser></c:barChart></x:barChart>
                <c:lineChart></c:lineChart>
            </c:plotArea></c:chart>
            <c:style val="4"/>
        </c:chartSpace>"#;

        let chart = parse_chart(xml.as_slice()).unwrap();
        let Some(TitleText::Reference(title)) = chart.title else {
            panic!("expected a chart title reference");
        };
        assert_eq!(title.formula, "Sheet1!$A$1");
        assert_eq!(chart.style, Some(4));
        assert!(matches!(
            chart.plot_area.type_groups.as_slice(),
            [TypeGroup::Line(_)]
        ));
    }

    #[test]
    fn rejects_non_chart_roots_and_trailing_roots() {
        let chart_namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        let foreign_root = br#"<x:chartSpace xmlns:x="urn:example"><x:chart/></x:chartSpace>"#;
        let trailing_root = format!(
            r#"<c:chartSpace xmlns:c="{chart_namespace}"><c:chart/></c:chartSpace><c:chartSpace xmlns:c="{chart_namespace}"/>"#
        );

        assert!(parse_chart(foreign_root.as_slice()).is_err());
        assert!(parse_chart(trailing_root.as_bytes()).is_err());
    }

    #[test]
    fn preserves_automatic_display_units_labels() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:valAx><c:axId val="1"/><c:scaling/>
                <c:axPos val="l"/><c:crossAx val="2"/><c:dispUnits>
                    <c:builtInUnit val="thousands"/><c:dispUnitsLbl/>
                </c:dispUnits>
            </c:valAx></c:plotArea></c:chart>
        </c:chartSpace>"#;

        let chart = parse_chart(xml.as_slice()).unwrap();
        let Axis::Value(axis) = &chart.plot_area.axes[0] else {
            panic!("expected value axis");
        };
        let display_units = axis.display_units.as_ref().unwrap();
        assert_eq!(display_units.built_in_unit, Some(BuiltInUnit::Thousands));
        assert!(display_units.show_label);
        assert!(display_units.label.is_none());
        assert!(display_units.layout.is_none());

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        assert!(
            std::str::from_utf8(&output)
                .unwrap()
                .contains("<c:dispUnitsLbl>")
        );
    }

    #[test]
    fn writer_rejects_invalid_display_units_and_duplicate_legend_entries() {
        let mut chart = Chart::new();
        let mut axis = ValueAxis::new(1, AxisPosition::Left, 2);
        let mut units = DisplayUnits::custom(1_000.0);
        units.built_in_unit = Some(BuiltInUnit::Thousands);
        axis.display_units = Some(units);
        chart.plot_area.axes.push(Axis::Value(axis));
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &chart).is_err());

        let mut chart = Chart::new();
        let mut axis = ValueAxis::new(1, AxisPosition::Left, 2);
        axis.display_units = Some(DisplayUnits::custom(f64::NAN));
        chart.plot_area.axes.push(Axis::Value(axis));
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &chart).is_err());

        let mut chart = Chart::new();
        let legend = Legend {
            entries: vec![LegendEntry::new(4), LegendEntry::new(4)],
            ..Legend::default()
        };
        chart.legend = Some(legend);
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &chart).is_err());
    }

    #[test]
    fn rejects_truncated_and_invalid_chart_values() {
        for xml in [
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>"#.as_slice(),
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotVisOnly val="yes"/></c:chart></c:chartSpace>"#.as_slice(),
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:dispBlanksAs val="empty"/></c:chart></c:chartSpace>"#.as_slice(),
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:style val="0"/><c:chart></c:chart></c:chartSpace>"#.as_slice(),
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:view3D><c:perspective val="241"/></c:view3D></c:chart></c:chartSpace>"#.as_slice(),
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:legend><c:legendPos val="center"/></c:legend></c:chart></c:chartSpace>"#.as_slice(),
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
        bubble.bubble_3d = true;
        bubble.size_represents = "w".to_string();
        let mut bubble_series = Series::new(0);
        bubble_series.x_values = Some(NumericData::from_values(vec![1.0]));
        bubble_series.y_values = Some(NumericData::from_values(vec![2.0]));
        bubble_series.bubble_sizes = Some(NumericData::from_values(vec![3.0]));
        bubble_series.bubble_3d = true;
        bubble.common.series.push(bubble_series);
        let mut scatter = ScatterTypeGroup::new(ScatterStyle::SmoothMarker);
        let mut scatter_series = Series::new(4);
        scatter_series.marker_symbol = Some(MarkerStyle::Star);
        scatter_series.marker_size = Some(9);
        scatter_series.smooth = true;
        let mut point = DataPoint::new(2).with_marker(7, MarkerStyle::Diamond);
        point.invert_if_negative = true;
        point.bubble_3d = Some(false);
        point.explosion = Some(15);
        scatter_series.data_points.push(point);
        let mut labels = DataLabels::new()
            .with_position(DataLabelPosition::Top)
            .with_show_value(true);
        labels.number_format = Some(NumberFormat::new("0.0%").with_source_linked(false));
        labels.show_series_name = true;
        labels.show_leader_lines = true;
        labels.separator = Some(" & ".to_string());
        let mut point_label = DataLabel::new(2);
        point_label.layout = Some(Layout::new().with_position(0.6, 0.7));
        point_label.text = Some(TitleText::from_ref("Sheet1!$E$2"));
        point_label.number_format = Some(NumberFormat::new("$0.00"));
        point_label.position = Some(DataLabelPosition::Left);
        point_label.show_category_name = true;
        point_label.separator = Some(" / ".to_string());
        labels.labels.push(point_label);
        scatter_series.data_labels = Some(labels);
        let mut trendline = Trendline::linear();
        trendline.name = Some("Forecast & fit".to_string());
        trendline.forward = Some(2.5);
        trendline.intercept = Some(-1.0);
        trendline.display_equation = true;
        trendline.display_r_squared = true;
        trendline.show_label = true;
        trendline.label = Some(TitleText::from_ref("Sheet1!$F$2"));
        trendline.label_layout = Some(Layout::new().with_size(0.3, 0.2));
        trendline.label_number_format = Some(NumberFormat::new("0.000").with_source_linked(false));
        scatter_series.trendlines.push(trendline);
        scatter_series.error_bars.push(ErrorBar {
            direction: ErrorBarDirection::Y,
            error_type: ErrorBarType::Both,
            value_type: ErrorBarValueType::Fixed,
            value: Some(1.5),
            plus_values: None,
            minus_values: None,
            no_end_cap: true,
        });
        scatter_series.error_bars.push(ErrorBar {
            direction: ErrorBarDirection::X,
            error_type: ErrorBarType::Plus,
            value_type: ErrorBarValueType::Custom,
            value: None,
            plus_values: Some(NumericData::from_ref("Sheet1!$D$2:$D$4")),
            minus_values: None,
            no_end_cap: false,
        });
        scatter.common.series.push(scatter_series);

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
            TypeGroup::Scatter(scatter),
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
        assert!(group.bubble_3d);
        assert_eq!(group.size_represents, "w");
        assert_eq!(
            group.common.series[0].x_values.as_ref().unwrap().values,
            [1.0]
        );
        assert_eq!(
            group.common.series[0].y_values.as_ref().unwrap().values,
            [2.0]
        );
        assert_eq!(
            group.common.series[0].bubble_sizes.as_ref().unwrap().values,
            [3.0]
        );
        assert!(group.common.series[0].bubble_3d);
        let TypeGroup::Scatter(group) = &parsed.plot_area.type_groups[11] else {
            unreachable!();
        };
        let series = &group.common.series[0];
        assert_eq!(series.marker_symbol, Some(MarkerStyle::Star));
        assert_eq!(series.marker_size, Some(9));
        assert!(series.smooth);
        assert_eq!(series.data_points.len(), 1);
        assert_eq!(series.data_points[0].index, 2);
        assert_eq!(series.data_points[0].marker_size, Some(7));
        assert_eq!(
            series.data_points[0].marker_symbol,
            Some(MarkerStyle::Diamond)
        );
        assert!(series.data_points[0].invert_if_negative);
        assert_eq!(series.data_points[0].bubble_3d, Some(false));
        assert_eq!(series.data_points[0].explosion, Some(15));
        let labels = series.data_labels.as_ref().unwrap();
        assert_eq!(labels.position, Some(DataLabelPosition::Top));
        assert!(labels.show_value);
        assert!(labels.show_series_name);
        assert!(labels.show_leader_lines);
        let number_format = labels.number_format.as_ref().unwrap();
        assert_eq!(number_format.format_code, "0.0%");
        assert!(!number_format.source_linked);
        assert_eq!(labels.separator.as_deref(), Some(" & "));
        assert_eq!(labels.labels.len(), 1);
        let point_label = &labels.labels[0];
        assert_eq!(point_label.index, 2);
        assert_eq!(point_label.layout.as_ref().unwrap().x, Some(0.6));
        assert_eq!(point_label.layout.as_ref().unwrap().y, Some(0.7));
        let Some(TitleText::Reference(text)) = point_label.text.as_ref() else {
            panic!("expected point data-label formula");
        };
        assert_eq!(text.formula, "Sheet1!$E$2");
        assert_eq!(
            point_label.number_format.as_ref().unwrap().format_code,
            "$0.00"
        );
        assert_eq!(point_label.position, Some(DataLabelPosition::Left));
        assert!(point_label.show_category_name);
        assert_eq!(point_label.separator.as_deref(), Some(" / "));
        assert_eq!(series.trendlines.len(), 1);
        assert_eq!(series.trendlines[0].name.as_deref(), Some("Forecast & fit"));
        assert_eq!(series.trendlines[0].forward, Some(2.5));
        assert_eq!(series.trendlines[0].intercept, Some(-1.0));
        assert!(series.trendlines[0].display_equation);
        assert!(series.trendlines[0].display_r_squared);
        assert!(series.trendlines[0].show_label);
        let Some(TitleText::Reference(label)) = series.trendlines[0].label.as_ref() else {
            panic!("expected trendline-label formula");
        };
        assert_eq!(label.formula, "Sheet1!$F$2");
        assert_eq!(
            series.trendlines[0].label_layout.as_ref().unwrap().width,
            Some(0.3)
        );
        assert_eq!(
            series.trendlines[0].label_layout.as_ref().unwrap().height,
            Some(0.2)
        );
        let number_format = series.trendlines[0].label_number_format.as_ref().unwrap();
        assert_eq!(number_format.format_code, "0.000");
        assert!(!number_format.source_linked);
        assert_eq!(series.error_bars.len(), 2);
        assert_eq!(series.error_bars[0].direction, ErrorBarDirection::Y);
        assert_eq!(series.error_bars[0].value, Some(1.5));
        assert!(series.error_bars[0].no_end_cap);
        assert_eq!(series.error_bars[1].value_type, ErrorBarValueType::Custom);
        assert_eq!(
            series.error_bars[1]
                .plus_values
                .as_ref()
                .unwrap()
                .source_ref
                .as_ref()
                .unwrap()
                .formula,
            "Sheet1!$D$2:$D$4"
        );
    }

    #[test]
    fn writer_round_trips_modeled_axis_properties_in_one_scaling_block() {
        let mut category = CategoryAxis::new(10, AxisPosition::Bottom, 20);
        category.common.orientation = AxisOrientation::MaxMin;
        category.common.title = Some(TitleText::from_string("Quarter"));
        category.common.title_overlay = true;
        category.common.layout = Some(Layout::new().with_position(0.3, 0.4));
        category.common.number_format =
            Some(NumberFormat::new("mmm-yy \"fiscal\"").with_source_linked(false));
        category.common.major_tick_mark = TickMark::Cross;
        category.common.minor_tick_mark = TickMark::In;
        category.common.tick_label_position = TickLabelPosition::Low;
        category.common.deleted = true;
        category.common.cross_mode = AxisCrossMode::Max;
        category.common.show_major_gridlines = true;
        category.auto = false;
        category.label_align = Some(AxisLabelAlign::Right);
        category.label_offset = Some(250);
        category.tick_label_skip = Some(2);
        category.tick_mark_skip = Some(3);
        category.no_multi_level = true;

        let mut value = ValueAxis::new(20, AxisPosition::Left, 10);
        value.common.crosses_at = Some(0.5);
        value.common.show_minor_gridlines = true;
        value.min = Some(-5.0);
        value.max = Some(100.0);
        value.major_unit = Some(10.0);
        value.minor_unit = Some(2.0);
        value.log_base = Some(10.0);
        value.cross_between = AxisCrossBetween::MidCategory;
        let mut display_units = DisplayUnits::built_in(BuiltInUnit::Millions);
        display_units.show_label = true;
        display_units.label = Some(TitleText::from_string("Millions sold"));
        display_units.layout = Some(Layout::new().with_position(0.15, 0.25));
        value.display_units = Some(display_units);

        let mut date = DateAxis::new(30, AxisPosition::Top, 40);
        date.min = Some(45_000.0);
        date.max = Some(46_000.0);
        date.major_unit = Some(2.0);
        date.minor_unit = Some(1.0);
        date.major_time_unit = Some(TimeUnit::Months);
        date.minor_time_unit = Some(TimeUnit::Days);
        date.base_time_unit = Some(TimeUnit::Years);
        date.auto = false;

        let mut series = SeriesAxis::new(40, AxisPosition::Right, 30);
        series.tick_label_skip = Some(4);
        series.tick_mark_skip = Some(5);

        let mut chart = Chart::new();
        chart.title = Some(TitleText::from_ref("Sheet1!$C$1"));
        chart.title_layout = Some(Layout::new().with_size(0.5, 0.1));
        chart.title_overlay = true;
        let mut layout = Layout::new().with_position(0.1, 0.2).with_size(0.7, 0.6);
        layout.target = Some(LayoutTarget::Inner);
        layout.x_mode = Some(LayoutMode::Factor);
        layout.y_mode = Some(LayoutMode::Edge);
        chart.plot_area.layout = Some(layout);
        let mut legend = Legend::new(LegendPosition::Top).with_overlay(true);
        legend.layout = Some(Layout::new().with_size(0.4, 0.2));
        let mut deleted_entry = LegendEntry::new(2);
        deleted_entry.deleted = true;
        legend.entries = vec![deleted_entry, LegendEntry::new(3)];
        chart.legend = Some(legend);
        chart.plot_area.axes = vec![
            Axis::Category(category),
            Axis::Value(value),
            Axis::Date(date),
            Axis::Series(series),
        ];

        let mut xml = Vec::new();
        crate::charts::writer::write_chart(&mut xml, &chart).unwrap();
        let xml_text = std::str::from_utf8(&xml).unwrap();
        assert_eq!(xml_text.matches("<c:scaling>").count(), 4);
        let parsed = parse_chart(xml.as_slice()).unwrap();
        let Some(TitleText::Reference(title)) = parsed.title.as_ref() else {
            panic!("expected chart title reference");
        };
        assert_eq!(title.formula, "Sheet1!$C$1");
        assert!(parsed.title_overlay);
        assert_eq!(parsed.title_layout.as_ref().unwrap().width, Some(0.5));
        assert_eq!(parsed.title_layout.as_ref().unwrap().height, Some(0.1));
        let layout = parsed.plot_area.layout.as_ref().unwrap();
        assert_eq!(layout.x, Some(0.1));
        assert_eq!(layout.y, Some(0.2));
        assert_eq!(layout.width, Some(0.7));
        assert_eq!(layout.height, Some(0.6));
        assert_eq!(layout.target, Some(LayoutTarget::Inner));
        assert_eq!(layout.x_mode, Some(LayoutMode::Factor));
        assert_eq!(layout.y_mode, Some(LayoutMode::Edge));
        assert_eq!(parsed.plot_area.axes.len(), 4);

        let Axis::Category(category) = &parsed.plot_area.axes[0] else {
            unreachable!();
        };
        assert_eq!(category.common.orientation, AxisOrientation::MaxMin);
        assert!(category.common.deleted);
        assert_eq!(category.common.major_tick_mark, TickMark::Cross);
        assert_eq!(category.common.minor_tick_mark, TickMark::In);
        assert_eq!(category.common.tick_label_position, TickLabelPosition::Low);
        assert_eq!(category.common.cross_mode, AxisCrossMode::Max);
        assert!(category.common.show_major_gridlines);
        let Some(TitleText::Literal(title)) = category.common.title.as_ref() else {
            panic!("expected literal category-axis title");
        };
        assert_eq!(title.text, "Quarter");
        assert!(category.common.title_overlay);
        assert_eq!(category.common.layout.as_ref().unwrap().x, Some(0.3));
        assert_eq!(category.common.layout.as_ref().unwrap().y, Some(0.4));
        assert_eq!(category.label_align, Some(AxisLabelAlign::Right));
        assert_eq!(category.label_offset, Some(250));
        assert_eq!(category.tick_label_skip, Some(2));
        assert_eq!(category.tick_mark_skip, Some(3));
        assert!(category.no_multi_level);
        let number_format = category.common.number_format.as_ref().unwrap();
        assert_eq!(number_format.format_code, "mmm-yy \"fiscal\"");
        assert!(!number_format.source_linked);

        let Axis::Value(value) = &parsed.plot_area.axes[1] else {
            unreachable!();
        };
        assert_eq!(value.min, Some(-5.0));
        assert_eq!(value.max, Some(100.0));
        assert_eq!(value.log_base, Some(10.0));
        assert_eq!(value.major_unit, Some(10.0));
        assert_eq!(value.minor_unit, Some(2.0));
        assert_eq!(value.common.crosses_at, Some(0.5));
        assert!(value.common.show_minor_gridlines);
        assert_eq!(value.cross_between, AxisCrossBetween::MidCategory);
        assert_eq!(
            value.display_units.as_ref().unwrap().built_in_unit,
            Some(BuiltInUnit::Millions)
        );
        let display_units = value.display_units.as_ref().unwrap();
        assert!(display_units.show_label);
        let Some(TitleText::Literal(label)) = display_units.label.as_ref() else {
            panic!("expected literal display-units label");
        };
        assert_eq!(label.text, "Millions sold");
        assert_eq!(display_units.layout.as_ref().unwrap().x, Some(0.15));
        assert_eq!(display_units.layout.as_ref().unwrap().y, Some(0.25));

        let legend = parsed.legend.as_ref().unwrap();
        assert_eq!(legend.position, LegendPosition::Top);
        assert!(legend.overlay);
        assert_eq!(legend.layout.as_ref().unwrap().width, Some(0.4));
        assert_eq!(legend.layout.as_ref().unwrap().height, Some(0.2));
        assert_eq!(legend.entries.len(), 2);
        assert_eq!(legend.entries[0].index, 2);
        assert!(legend.entries[0].deleted);
        assert_eq!(legend.entries[1].index, 3);
        assert!(!legend.entries[1].deleted);

        let Axis::Date(date) = &parsed.plot_area.axes[2] else {
            unreachable!();
        };
        assert_eq!(date.min, Some(45_000.0));
        assert_eq!(date.max, Some(46_000.0));
        assert_eq!(date.major_time_unit, Some(TimeUnit::Months));
        assert_eq!(date.minor_time_unit, Some(TimeUnit::Days));
        assert_eq!(date.base_time_unit, Some(TimeUnit::Years));
        assert!(!date.auto);

        let Axis::Series(series) = &parsed.plot_area.axes[3] else {
            unreachable!();
        };
        assert_eq!(series.tick_label_skip, Some(4));
        assert_eq!(series.tick_mark_skip, Some(5));
    }
}
