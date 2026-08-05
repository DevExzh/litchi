//! Chart XML reader.
//!
//! This module provides functionality to parse chart XML files
//! from OOXML packages.

use super::model::*;

use crate::chart::axis::{
    Axis, AxisCommon, AxisCrossBetween, AxisCrossMode, AxisLabelAlign, BuiltInUnit, CategoryAxis,
    DateAxis, DisplayUnits, SeriesAxis, TimeUnit, ValueAxis,
};
use crate::chart::bubble::{Scale as BubbleScale, Size as BubbleSize};
use crate::chart::data::{
    DataSourceRef, Layout, NumberFormat, NumericData, RichText, StringData, TitleText,
};
use crate::chart::legend::{Legend, LegendEntry};
use crate::chart::model::{
    Chart, ColorMapOverride, ColorMapping, ColorSchemeIndex, ExtensionList, ExternalData,
    HeaderFooter, PageMargins, PageOrientation, PageSetup, PictureFormat, PictureOptions,
    PivotFormat, PivotSource, PrintSettings, Protection, ShapeProperties, TextProperties,
    UserShapes, View3D, WallFloor,
};
use crate::chart::plot_area::{
    Area3DTypeGroup, AreaTypeGroup, BandFormat, Bar3DTypeGroup, BarShape, BarTypeGroup,
    BubbleTypeGroup, DataTable, DoughnutTypeGroup, Line3DTypeGroup, LineTypeGroup, Lines,
    OfPieTypeGroup, Pie3DTypeGroup, PieTypeGroup, PlotArea, RadarTypeGroup, ScatterTypeGroup,
    StockTypeGroup, Surface3DTypeGroup, SurfaceTypeGroup, TypeGroup, TypeGroupCommon, UpDownBars,
};
use crate::chart::series::{
    DataLabel, DataLabels, DataPoint, ErrorBar, ErrorBarDirection, ErrorBarType, ErrorBarValueType,
    Marker, Series, Trendline, TrendlineType,
};
use crate::chart::types::{
    AxisOrientation, AxisPosition, BarDirection, BarGrouping, DataLabelPosition, DisplayBlanks,
    LayoutMode, LayoutTarget, LegendPosition, MarkerStyle, OfPieSplitType, OfPieType, RadarStyle,
    ScatterStyle, TickLabelPosition, TickMark,
};
use crate::{Error, Result};
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use std::io::{BufRead, Read};

/// Parse a chart XML document.
pub fn read<R: BufRead>(reader: R) -> Result<Chart> {
    let limits = litchi_ooxml_common::mce::Limits::default();
    let mut input = Vec::new();
    reader
        .take((limits.max_input_bytes as u64).saturating_add(1))
        .read_to_end(&mut input)?;
    if input.len() > limits.max_input_bytes {
        return Err(Error::Mce(litchi_ooxml_common::mce::Error::LimitExceeded(
            "input bytes".into(),
        )));
    }
    let xml = litchi_ooxml_common::mce::process_markup_compatibility(
        &input,
        &litchi_ooxml_common::mce::Capabilities::default(),
        &limits,
    )?
    .xml
    .into_owned();
    let mut xml_reader = ChartXmlReader::from_reader(std::io::Cursor::new(xml));
    xml_reader.config_mut().trim_text(false);

    let mut chart = Chart::new();
    let mut buf = Vec::new();
    let mut saw_chart = false;
    let mut closed_chart = false;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lang" => {
                if chart.language.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate language declarations".into(),
                    ));
                }
                chart.language = Some(required_string_attr(
                    e,
                    b"val",
                    xml_reader.decoder(),
                    "chart language",
                )?);
                consume_empty_chart_element(&mut xml_reader, b"lang", "chart language")?;
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"lang" => {
                if chart.language.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate language declarations".into(),
                    ));
                }
                chart.language = Some(required_string_attr(
                    e,
                    b"val",
                    xml_reader.decoder(),
                    "chart language",
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pivotSource" => {
                if chart.pivot_source.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate pivot sources".into(),
                    ));
                }
                chart.pivot_source = Some(parse_pivot_source(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"pivotSource" => {
                return Err(Error::Invalid(
                    "chart pivot source requires a name and format ID".into(),
                ));
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrMapOvr" => {
                if chart.color_map_override.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate color-map overrides".into(),
                    ));
                }
                chart.color_map_override = Some(parse_color_map_override(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"clrMapOvr" => {
                return Err(Error::Invalid(
                    "chart color-map override requires a mapping choice".into(),
                ));
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"protection" => {
                if chart.protection.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate protection settings".into(),
                    ));
                }
                chart.protection = Some(parse_chart_protection(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"protection" => {
                if chart.protection.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate protection settings".into(),
                    ));
                }
                chart.protection = Some(Protection::default());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pivotFmts" => {
                if chart.pivot_formats.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate pivot-format collections".into(),
                    ));
                }
                chart.pivot_formats = Some(parse_pivot_formats(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"pivotFmts" => {
                if chart.pivot_formats.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate pivot-format collections".into(),
                    ));
                }
                chart.pivot_formats = Some(Vec::new());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"view3D" => {
                if chart.view_3d.is_some() {
                    return Err(Error::Invalid("chart contains duplicate 3D views".into()));
                }
                chart.view_3d = Some(parse_view_3d(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"view3D" => {
                if chart.view_3d.is_some() {
                    return Err(Error::Invalid("chart contains duplicate 3D views".into()));
                }
                chart.view_3d = Some(View3D::new());
            },
            Ok(Event::Start(ref e))
                if matches!(
                    e.local_name().as_ref(),
                    b"floor" | b"backWall" | b"sideWall"
                ) =>
            {
                let target = match e.local_name().as_ref() {
                    b"floor" => &mut chart.floor,
                    b"backWall" => &mut chart.back_wall,
                    b"sideWall" => &mut chart.side_wall,
                    _ => {
                        return Err(Error::Invalid("invalid chart 3D surface element".into()));
                    },
                };
                if target.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate 3D surfaces".into(),
                    ));
                }
                *target = Some(parse_wall_floor(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e))
                if matches!(
                    e.local_name().as_ref(),
                    b"floor" | b"backWall" | b"sideWall"
                ) =>
            {
                let target = match e.local_name().as_ref() {
                    b"floor" => &mut chart.floor,
                    b"backWall" => &mut chart.back_wall,
                    b"sideWall" => &mut chart.side_wall,
                    _ => {
                        return Err(Error::Invalid("invalid chart 3D surface element".into()));
                    },
                };
                if target.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate 3D surfaces".into(),
                    ));
                }
                *target = Some(WallFloor::new());
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"plotArea" => {
                chart.plot_area = PlotArea::new();
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"printSettings" => {
                if chart.print_settings.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate print settings".into(),
                    ));
                }
                chart.print_settings = Some(parse_print_settings(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"printSettings" => {
                if chart.print_settings.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate print settings".into(),
                    ));
                }
                chart.print_settings = Some(PrintSettings::new());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"externalData" => {
                if chart.external_data.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate external-data relationships".into(),
                    ));
                }
                let relationship_id = required_chart_relationship_id(&xml_reader, e)?;
                chart.external_data = Some(parse_external_data(&mut xml_reader, relationship_id)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"externalData" => {
                if chart.external_data.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate external-data relationships".into(),
                    ));
                }
                chart.external_data = Some(ExternalData::new(required_chart_relationship_id(
                    &xml_reader,
                    e,
                )?));
            },
            Ok(Event::Start(ref e)) if closed_chart && e.local_name().as_ref() == b"spPr" => {
                if chart.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate chart-space shape properties".into(),
                    ));
                }
                chart.shape_properties = Some(ShapeProperties::from_xml(
                    xml_reader.capture_fragment(e, "chart-space shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if closed_chart && e.local_name().as_ref() == b"spPr" => {
                if chart.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate chart-space shape properties".into(),
                    ));
                }
                chart.shape_properties = Some(ShapeProperties::from_xml(
                    xml_reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if closed_chart && e.local_name().as_ref() == b"txPr" => {
                if chart.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate chart-space text properties".into(),
                    ));
                }
                chart.text_properties = Some(TextProperties::from_xml(
                    xml_reader.capture_fragment(e, "chart-space text properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if closed_chart && e.local_name().as_ref() == b"txPr" => {
                if chart.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate chart-space text properties".into(),
                    ));
                }
                chart.text_properties = Some(TextProperties::from_xml(
                    xml_reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"userShapes" => {
                if chart.user_shapes.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate user-shapes relationships".into(),
                    ));
                }
                chart.user_shapes = Some(UserShapes::new(required_chart_relationship_id(
                    &xml_reader,
                    e,
                )?));
                consume_empty_chart_element(
                    &mut xml_reader,
                    b"userShapes",
                    "chart user-shapes relationship",
                )?;
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"userShapes" => {
                if chart.user_shapes.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate user-shapes relationships".into(),
                    ));
                }
                chart.user_shapes = Some(UserShapes::new(required_chart_relationship_id(
                    &xml_reader,
                    e,
                )?));
            },
            Ok(Event::Start(ref e))
                if saw_chart && !closed_chart && e.local_name().as_ref() == b"extLst" =>
            {
                if chart.chart_extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate chart extension lists".into(),
                    ));
                }
                chart.chart_extension_list = Some(ExtensionList::from_xml(
                    xml_reader.capture_fragment(e, "chart extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e))
                if saw_chart && !closed_chart && e.local_name().as_ref() == b"extLst" =>
            {
                if chart.chart_extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate chart extension lists".into(),
                    ));
                }
                chart.chart_extension_list = Some(ExtensionList::from_xml(
                    xml_reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if closed_chart && e.local_name().as_ref() == b"extLst" => {
                if chart.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate chart-space extension lists".into(),
                    ));
                }
                chart.extension_list = Some(ExtensionList::from_xml(
                    xml_reader.capture_fragment(e, "chart-space extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if closed_chart && e.local_name().as_ref() == b"extLst" => {
                if chart.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart contains duplicate chart-space extension lists".into(),
                    ));
                }
                chart.extension_list = Some(ExtensionList::from_xml(
                    xml_reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"chart" => saw_chart = true,
                    b"chartSpace" => {},
                    b"title" => {
                        if chart.title.is_some() {
                            return Err(Error::Invalid("chart contains duplicate titles".into()));
                        }
                        let title = parse_title(&mut xml_reader)?;
                        chart.title = Some(title.text);
                        chart.title_layout = title.layout;
                        chart.title_overlay = title.overlay;
                        chart.title_shape_properties = title.shape_properties;
                        chart.title_text_properties = title.text_properties;
                        chart.title_extension_list = title.extension_list;
                    },
                    b"autoTitleDeleted" => {
                        chart.auto_title_deleted = parse_bool_attr(e)?;
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
                return Err(Error::Invalid(
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

fn parse_pivot_source<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<PivotSource> {
    let mut name = None;
    let mut format_id = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"name" => {
                if name.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot source contains duplicate names".into(),
                    ));
                }
                name = Some(parse_text_element(reader, b"name")?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"name" => {
                if name.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot source contains duplicate names".into(),
                    ));
                }
                name = Some(String::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"fmtId" => {
                if format_id.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot source contains duplicate format IDs".into(),
                    ));
                }
                format_id = Some(required_u32_attr(element, "chart pivot-source format ID")?);
                consume_empty_chart_element(reader, b"fmtId", "chart pivot-source format ID")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"fmtId" => {
                if format_id.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot source contains duplicate format IDs".into(),
                    ));
                }
                format_id = Some(required_u32_attr(element, "chart pivot-source format ID")?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"pivotSource" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart pivot source".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(PivotSource::new(
        name.ok_or_else(|| missing_attribute("chart pivot-source name"))?,
        format_id.ok_or_else(|| missing_attribute("chart pivot-source format ID"))?,
    ))
}

fn parse_chart_protection<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Protection> {
    let mut protection = Protection::default();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) => {
                let field = match element.local_name().as_ref() {
                    b"chartObject" => &mut protection.chart_object,
                    b"data" => &mut protection.data,
                    b"formatting" => &mut protection.formatting,
                    b"selection" => &mut protection.selection,
                    b"userInterface" => &mut protection.user_interface,
                    _ => {
                        buf.clear();
                        continue;
                    },
                };
                if field.is_some() {
                    return Err(Error::Invalid(format!(
                        "chart protection contains duplicate {} settings",
                        String::from_utf8_lossy(element.local_name().as_ref())
                    )));
                }
                *field = Some(parse_bool_attr(element)?);
                let end_name = element.local_name().as_ref().to_vec();
                consume_empty_chart_element(reader, &end_name, "chart protection switch")?;
            },
            Ok(Event::Empty(ref element)) => {
                let field = match element.local_name().as_ref() {
                    b"chartObject" => &mut protection.chart_object,
                    b"data" => &mut protection.data,
                    b"formatting" => &mut protection.formatting,
                    b"selection" => &mut protection.selection,
                    b"userInterface" => &mut protection.user_interface,
                    _ => {
                        buf.clear();
                        continue;
                    },
                };
                if field.is_some() {
                    return Err(Error::Invalid(format!(
                        "chart protection contains duplicate {} settings",
                        String::from_utf8_lossy(element.local_name().as_ref())
                    )));
                }
                *field = Some(parse_bool_attr(element)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"protection" => {
                return Ok(protection);
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid("chart protection is not closed".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
}

fn parse_color_map_override<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<ColorMapOverride> {
    let mut mapping = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element))
                if element.local_name().as_ref() == b"masterClrMapping" =>
            {
                if mapping.is_some() {
                    return Err(Error::Invalid(
                        "chart color-map override contains multiple choices".into(),
                    ));
                }
                consume_empty_chart_element(
                    reader,
                    b"masterClrMapping",
                    "chart master color mapping",
                )?;
                mapping = Some(ColorMapOverride::Master);
            },
            Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"masterClrMapping" =>
            {
                set_color_map_override_choice(&mut mapping, ColorMapOverride::Master)?;
            },
            Ok(Event::Start(ref element))
                if element.local_name().as_ref() == b"overrideClrMapping" =>
            {
                if mapping.is_some() {
                    return Err(Error::Invalid(
                        "chart color-map override contains multiple choices".into(),
                    ));
                }
                let value = parse_color_mapping(element)?;
                consume_empty_chart_element(
                    reader,
                    b"overrideClrMapping",
                    "chart override color mapping",
                )?;
                mapping = Some(ColorMapOverride::Override(value));
            },
            Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"overrideClrMapping" =>
            {
                let value = ColorMapOverride::Override(parse_color_mapping(element)?);
                set_color_map_override_choice(&mut mapping, value)?;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() != IGNORED_NAMESPACE_ELEMENT.as_bytes() =>
            {
                return Err(Error::Invalid(
                    "chart color-map override contains an unexpected choice".into(),
                ));
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"clrMapOvr" => {
                return mapping.ok_or_else(|| {
                    Error::Invalid("chart color-map override requires a mapping choice".into())
                });
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
                    "chart color-map override is not closed".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
}

fn set_color_map_override_choice(
    target: &mut Option<ColorMapOverride>,
    value: ColorMapOverride,
) -> Result<()> {
    if target.is_some() {
        return Err(Error::Invalid(
            "chart color-map override contains multiple choices".into(),
        ));
    }
    *target = Some(value);
    Ok(())
}

fn parse_color_mapping(element: &BytesStart<'_>) -> Result<ColorMapping> {
    Ok(ColorMapping {
        background1: required_color_scheme_index(element, b"bg1")?,
        text1: required_color_scheme_index(element, b"tx1")?,
        background2: required_color_scheme_index(element, b"bg2")?,
        text2: required_color_scheme_index(element, b"tx2")?,
        accent1: required_color_scheme_index(element, b"accent1")?,
        accent2: required_color_scheme_index(element, b"accent2")?,
        accent3: required_color_scheme_index(element, b"accent3")?,
        accent4: required_color_scheme_index(element, b"accent4")?,
        accent5: required_color_scheme_index(element, b"accent5")?,
        accent6: required_color_scheme_index(element, b"accent6")?,
        hyperlink: required_color_scheme_index(element, b"hlink")?,
        followed_hyperlink: required_color_scheme_index(element, b"folHlink")?,
    })
}

fn required_chart_relationship_id<R: BufRead>(
    reader: &ChartXmlReader<R>,
    element: &BytesStart<'_>,
) -> Result<String> {
    reader
        .relationship_attribute_value(element, b"id")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid("chart relationship ID is required".into()))
}

fn parse_external_data<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    relationship_id: String,
) -> Result<ExternalData> {
    let mut external_data = ExternalData::new(relationship_id);
    let mut saw_auto_update = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"autoUpdate" => {
                if saw_auto_update {
                    return Err(Error::Invalid(
                        "chart external data contains duplicate auto-update settings".into(),
                    ));
                }
                saw_auto_update = true;
                external_data.auto_update = Some(parse_bool_attr(element)?);
                consume_empty_chart_element(
                    reader,
                    b"autoUpdate",
                    "chart external-data auto-update setting",
                )?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"autoUpdate" => {
                if saw_auto_update {
                    return Err(Error::Invalid(
                        "chart external data contains duplicate auto-update settings".into(),
                    ));
                }
                saw_auto_update = true;
                external_data.auto_update = Some(parse_bool_attr(element)?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() != IGNORED_NAMESPACE_ELEMENT.as_bytes() =>
            {
                return Err(Error::Invalid(
                    "chart external data contains an unexpected child".into(),
                ));
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"externalData" => {
                return Ok(external_data);
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid("chart external data is not closed".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
}

fn required_color_scheme_index(
    element: &BytesStart<'_>,
    attribute: &[u8],
) -> Result<ColorSchemeIndex> {
    let value = get_attr(element, attribute).ok_or_else(|| {
        Error::Invalid(format!(
            "chart color mapping is missing its {} assignment",
            String::from_utf8_lossy(attribute)
        ))
    })?;
    match value.as_slice() {
        b"dk1" => Ok(ColorSchemeIndex::Dark1),
        b"lt1" => Ok(ColorSchemeIndex::Light1),
        b"dk2" => Ok(ColorSchemeIndex::Dark2),
        b"lt2" => Ok(ColorSchemeIndex::Light2),
        b"accent1" => Ok(ColorSchemeIndex::Accent1),
        b"accent2" => Ok(ColorSchemeIndex::Accent2),
        b"accent3" => Ok(ColorSchemeIndex::Accent3),
        b"accent4" => Ok(ColorSchemeIndex::Accent4),
        b"accent5" => Ok(ColorSchemeIndex::Accent5),
        b"accent6" => Ok(ColorSchemeIndex::Accent6),
        b"hlink" => Ok(ColorSchemeIndex::Hyperlink),
        b"folHlink" => Ok(ColorSchemeIndex::FollowedHyperlink),
        _ => Err(invalid_attribute("chart color-scheme index", &value)),
    }
}

fn parse_pivot_formats<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Vec<PivotFormat>> {
    let mut formats = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"pivotFmt" => {
                let format = parse_pivot_format(reader)?;
                if formats
                    .iter()
                    .any(|existing: &PivotFormat| existing.index == format.index)
                {
                    return Err(Error::Invalid(format!(
                        "chart contains duplicate pivot-format index {}",
                        format.index
                    )));
                }
                formats.push(format);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"pivotFmt" => {
                return Err(Error::Invalid(
                    "chart pivot format is missing its index".into(),
                ));
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"pivotFmts" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart pivot formats".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(formats)
}

fn parse_pivot_format<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<PivotFormat> {
    let mut index = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut marker = None;
    let mut data_label = None;
    let mut extension_list = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart pivot-format shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart pivot-format text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"marker" => {
                if marker.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate markers".into(),
                    ));
                }
                marker = Some(parse_series_marker(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"marker" => {
                if marker.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate markers".into(),
                    ));
                }
                marker = Some(Marker::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dLbl" => {
                if data_label.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate data labels".into(),
                    ));
                }
                data_label = Some(parse_data_label(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dLbl" => {
                return Err(Error::Invalid(
                    "chart pivot-format data label is missing its index".into(),
                ));
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"idx" =>
            {
                if index.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate indexes".into(),
                    ));
                }
                index = Some(required_u32_attr(element, "chart pivot-format index")?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart pivot-format extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart pivot format contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"pivotFmt" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart pivot format".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut format =
        PivotFormat::new(index.ok_or_else(|| missing_attribute("chart pivot-format index"))?);
    format.shape_properties = shape_properties;
    format.text_properties = text_properties;
    format.marker = marker;
    format.data_label = data_label;
    format.extension_list = extension_list;
    Ok(format)
}

fn parse_print_settings<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<PrintSettings> {
    let mut settings = PrintSettings::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"headerFooter" => {
                if settings.header_footer.is_some() {
                    return Err(Error::Invalid(
                        "chart print settings contain duplicate headers and footers".into(),
                    ));
                }
                settings.header_footer = Some(parse_chart_header_footer(reader, element)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"headerFooter" => {
                if settings.header_footer.is_some() {
                    return Err(Error::Invalid(
                        "chart print settings contain duplicate headers and footers".into(),
                    ));
                }
                settings.header_footer = Some(parse_chart_header_footer_attributes(element)?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"pageMargins" => {
                if settings.page_margins.is_some() {
                    return Err(Error::Invalid(
                        "chart print settings contain duplicate page margins".into(),
                    ));
                }
                settings.page_margins = Some(parse_chart_page_margins(element)?);
                consume_empty_chart_element(reader, b"pageMargins", "chart page margins")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"pageMargins" => {
                if settings.page_margins.is_some() {
                    return Err(Error::Invalid(
                        "chart print settings contain duplicate page margins".into(),
                    ));
                }
                settings.page_margins = Some(parse_chart_page_margins(element)?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"pageSetup" => {
                if settings.page_setup.is_some() {
                    return Err(Error::Invalid(
                        "chart print settings contain duplicate page setup".into(),
                    ));
                }
                settings.page_setup = Some(parse_chart_page_setup(element)?);
                consume_empty_chart_element(reader, b"pageSetup", "chart page setup")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"pageSetup" => {
                if settings.page_setup.is_some() {
                    return Err(Error::Invalid(
                        "chart print settings contain duplicate page setup".into(),
                    ));
                }
                settings.page_setup = Some(parse_chart_page_setup(element)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"printSettings" => {
                break;
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart print settings".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(settings)
}

fn parse_chart_header_footer_attributes(element: &BytesStart<'_>) -> Result<HeaderFooter> {
    let mut header_footer = HeaderFooter::new();
    header_footer.align_with_margins = optional_bool_attr(
        element,
        b"alignWithMargins",
        true,
        "header/footer alignment",
    )?;
    header_footer.different_odd_even = optional_bool_attr(
        element,
        b"differentOddEven",
        false,
        "odd/even header/footer selection",
    )?;
    header_footer.different_first = optional_bool_attr(
        element,
        b"differentFirst",
        false,
        "first-page header/footer selection",
    )?;
    Ok(header_footer)
}

fn parse_chart_header_footer<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    element: &BytesStart<'_>,
) -> Result<HeaderFooter> {
    let mut header_footer = parse_chart_header_footer_attributes(element)?;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref child)) => {
                let value = match child.local_name().as_ref() {
                    b"oddHeader" => &mut header_footer.odd_header,
                    b"oddFooter" => &mut header_footer.odd_footer,
                    b"evenHeader" => &mut header_footer.even_header,
                    b"evenFooter" => &mut header_footer.even_footer,
                    b"firstHeader" => &mut header_footer.first_header,
                    b"firstFooter" => &mut header_footer.first_footer,
                    name if name == IGNORED_NAMESPACE_ELEMENT.as_bytes() => {
                        buf.clear();
                        continue;
                    },
                    _ => {
                        return Err(Error::Invalid(
                            "chart header/footer contains an unexpected child".into(),
                        ));
                    },
                };
                if value.is_some() {
                    return Err(Error::Invalid(
                        "chart header/footer contains a duplicate string".into(),
                    ));
                }
                *value = Some(parse_text_element(reader, child.local_name().as_ref())?);
            },
            Ok(Event::Empty(ref child)) => {
                let value = match child.local_name().as_ref() {
                    b"oddHeader" => &mut header_footer.odd_header,
                    b"oddFooter" => &mut header_footer.odd_footer,
                    b"evenHeader" => &mut header_footer.even_header,
                    b"evenFooter" => &mut header_footer.even_footer,
                    b"firstHeader" => &mut header_footer.first_header,
                    b"firstFooter" => &mut header_footer.first_footer,
                    name if name == IGNORED_NAMESPACE_ELEMENT.as_bytes() => {
                        buf.clear();
                        continue;
                    },
                    _ => {
                        return Err(Error::Invalid(
                            "chart header/footer contains an unexpected child".into(),
                        ));
                    },
                };
                if value.replace(String::new()).is_some() {
                    return Err(Error::Invalid(
                        "chart header/footer contains a duplicate string".into(),
                    ));
                }
            },
            Ok(Event::End(ref child)) if child.local_name().as_ref() == b"headerFooter" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart header/footer".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(header_footer)
}

fn parse_chart_page_margins(element: &BytesStart<'_>) -> Result<PageMargins> {
    Ok(PageMargins::new(
        required_named_f64_attr(element, b"l", "chart left page margin")?,
        required_named_f64_attr(element, b"r", "chart right page margin")?,
        required_named_f64_attr(element, b"t", "chart top page margin")?,
        required_named_f64_attr(element, b"b", "chart bottom page margin")?,
        required_named_f64_attr(element, b"header", "chart header page margin")?,
        required_named_f64_attr(element, b"footer", "chart footer page margin")?,
    ))
}

fn parse_chart_page_setup(element: &BytesStart<'_>) -> Result<PageSetup> {
    let mut setup = PageSetup::new();
    setup.paper_size = optional_u32_attr(element, b"paperSize", 1, "chart printer paper size")?;
    setup.first_page_number =
        optional_u32_attr(element, b"firstPageNumber", 1, "chart first page number")?;
    setup.orientation = match get_attr(element, b"orientation").as_deref() {
        None | Some(b"default") => PageOrientation::Default,
        Some(b"portrait") => PageOrientation::Portrait,
        Some(b"landscape") => PageOrientation::Landscape,
        Some(value) => return Err(invalid_attribute("chart page orientation", value)),
    };
    setup.black_and_white = optional_bool_attr(
        element,
        b"blackAndWhite",
        false,
        "chart black-and-white printing",
    )?;
    setup.draft = optional_bool_attr(element, b"draft", false, "chart draft printing")?;
    setup.use_first_page_number = optional_bool_attr(
        element,
        b"useFirstPageNumber",
        false,
        "chart first-page-number usage",
    )?;
    setup.horizontal_dpi = optional_i32_attr(
        element,
        b"horizontalDpi",
        600,
        "chart horizontal printer resolution",
    )?;
    setup.vertical_dpi = optional_i32_attr(
        element,
        b"verticalDpi",
        600,
        "chart vertical printer resolution",
    )?;
    setup.copies = optional_u32_attr(element, b"copies", 1, "chart print copies")?;
    Ok(setup)
}

fn consume_empty_chart_element<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    end_name: &[u8],
    description: &str,
) -> Result<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::End(ref element)) if element.local_name().as_ref() == end_name => {
                return Ok(());
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() != IGNORED_NAMESPACE_ELEMENT.as_bytes() =>
            {
                return Err(Error::Invalid(format!(
                    "{description} contains child elements"
                )));
            },
            Ok(Event::Text(ref text))
                if !text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim()
                    .is_empty() =>
            {
                return Err(Error::Invalid(format!("{description} contains text")));
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid(format!("unterminated {description}")));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
}

struct ParsedTitle {
    text: TitleText,
    layout: Option<Layout>,
    overlay: bool,
    shape_properties: Option<ShapeProperties>,
    text_properties: Option<TextProperties>,
    extension_list: Option<ExtensionList>,
}

fn parse_title<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<ParsedTitle> {
    let mut text = String::new();
    let mut formula = None;
    let mut layout = None;
    let mut overlay = false;
    let mut saw_overlay = false;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut extension_list = None;
    let mut buf = Vec::new();
    let mut in_text = false;
    let mut saw_text = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if layout.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate layouts".into(),
                    ));
                }
                layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                layout = Some(match layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(Error::Invalid(
                            "chart title contains duplicate layouts".into(),
                        ));
                    },
                });
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"overlay" =>
            {
                if saw_overlay {
                    return Err(Error::Invalid(
                        "chart title contains duplicate overlay flags".into(),
                    ));
                }
                overlay = parse_bool_attr(element)?;
                saw_overlay = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart title shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart title text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart title extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart title contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"f" => {
                if formula.is_some() {
                    return Err(Error::Invalid(
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
                text.push_str(&e.decode().map_err(|e| Error::Xml(e.to_string()))?);
            },
            Ok(Event::CData(e)) if in_text => {
                text.push_str(&e.decode().map_err(|e| Error::Xml(e.to_string()))?);
            },
            Ok(Event::GeneralRef(reference)) if in_text => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"t" => {
                in_text = false;
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"title" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let text = if let Some(formula) = formula {
        if saw_text {
            return Err(Error::Invalid(
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
        shape_properties,
        text_properties,
        extension_list,
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
                return Err(Error::Invalid("unterminated chart element".to_string()));
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
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"thickness" => {
                if wall_floor.thickness.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate thickness values".into(),
                    ));
                }
                wall_floor.thickness = Some(bounded_u32_attr(e, "chart wall thickness", 0, 4096)?);
                consume_empty_chart_element(reader, b"thickness", "chart surface thickness")?;
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"thickness" => {
                if wall_floor.thickness.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate thickness values".into(),
                    ));
                }
                wall_floor.thickness = Some(bounded_u32_attr(e, "chart wall thickness", 0, 4096)?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if wall_floor.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate shape properties".into(),
                    ));
                }
                wall_floor.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart surface shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if wall_floor.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate shape properties".into(),
                    ));
                }
                wall_floor.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pictureOptions" => {
                if wall_floor.picture_options.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate picture options".into(),
                    ));
                }
                wall_floor.picture_options = Some(parse_picture_options(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"pictureOptions" => {
                if wall_floor.picture_options.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate picture options".into(),
                    ));
                }
                wall_floor.picture_options = Some(PictureOptions::default());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if wall_floor.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate extension lists".into(),
                    ));
                }
                wall_floor.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(e, "chart surface extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if wall_floor.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart surface contains duplicate extension lists".into(),
                    ));
                }
                wall_floor.extension_list =
                    Some(ExtensionList::from_xml(reader.capture_empty_fragment(e)?)?);
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
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(wall_floor)
}

fn parse_picture_options<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<PictureOptions> {
    let mut options = PictureOptions::default();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref element))
                if is_picture_option_child(element.local_name().as_ref()) =>
            {
                parse_picture_option_child(&mut options, element, reader.decoder())?;
                let name = element.local_name().as_ref().to_vec();
                consume_empty_chart_element(reader, &name, "chart picture option")?;
            },
            Ok(Event::Empty(ref element))
                if is_picture_option_child(element.local_name().as_ref()) =>
            {
                parse_picture_option_child(&mut options, element, reader.decoder())?;
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"pictureOptions" => {
                break;
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart picture options".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buffer.clear();
    }
    Ok(options)
}

fn is_picture_option_child(name: &[u8]) -> bool {
    matches!(
        name,
        b"applyToFront" | b"applyToSides" | b"applyToEnd" | b"pictureFormat" | b"pictureStackUnit"
    )
}

fn parse_picture_option_child(
    options: &mut PictureOptions,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<()> {
    let target = match element.local_name().as_ref() {
        b"applyToFront" => Some(&mut options.apply_to_front),
        b"applyToSides" => Some(&mut options.apply_to_sides),
        b"applyToEnd" => Some(&mut options.apply_to_end),
        b"pictureFormat" => {
            if options.picture_format.is_some() {
                return Err(Error::Invalid(
                    "chart picture options contain duplicate formats".into(),
                ));
            }
            let value = required_string_attr(element, b"val", decoder, "chart picture format")?;
            options.picture_format = Some(match value.as_str() {
                "stretch" => PictureFormat::Stretch,
                "stack" => PictureFormat::Stack,
                "stackScale" => PictureFormat::StackScale,
                _ => {
                    return Err(Error::Invalid(format!(
                        "invalid chart picture format '{value}'"
                    )));
                },
            });
            None
        },
        b"pictureStackUnit" => {
            if options.picture_stack_unit.is_some() {
                return Err(Error::Invalid(
                    "chart picture options contain duplicate stack units".into(),
                ));
            }
            options.picture_stack_unit = Some(required_positive_f64_attr(
                element,
                "chart picture stack unit",
            )?);
            None
        },
        _ => {
            return Err(Error::Invalid(format!(
                "invalid chart picture option '{}'",
                String::from_utf8_lossy(element.local_name().as_ref())
            )));
        },
    };
    if let Some(target) = target {
        if target.is_some() {
            return Err(Error::Invalid(
                "chart picture options contain a duplicate apply flag".into(),
            ));
        }
        *target = Some(parse_bool_attr(element)?);
    }
    Ok(())
}

fn parse_plot_area<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<PlotArea> {
    let mut plot_area = PlotArea::new();
    let mut saw_data_table = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"layout" => {
                plot_area.layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"layout" => {
                plot_area.layout = Some(Layout::new());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dTable" => {
                if saw_data_table {
                    return Err(Error::Invalid(
                        "chart plot area contains duplicate data tables".into(),
                    ));
                }
                saw_data_table = true;
                plot_area.data_table = Some(parse_data_table(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dTable" => {
                if saw_data_table {
                    return Err(Error::Invalid(
                        "chart plot area contains duplicate data tables".into(),
                    ));
                }
                saw_data_table = true;
                plot_area.data_table = Some(DataTable::default());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if plot_area.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart plot area contains duplicate shape properties".into(),
                    ));
                }
                plot_area.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart plot-area shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if plot_area.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart plot area contains duplicate shape properties".into(),
                    ));
                }
                plot_area.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if plot_area.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart plot area contains duplicate extension lists".into(),
                    ));
                }
                plot_area.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(e, "chart plot-area extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if plot_area.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart plot area contains duplicate extension lists".into(),
                    ));
                }
                plot_area.extension_list =
                    Some(ExtensionList::from_xml(reader.capture_empty_fragment(e)?)?);
            },
            Ok(Event::Empty(ref e))
                if is_chart_type_group_name(e.local_name().as_ref())
                    || matches!(
                        e.local_name().as_ref(),
                        b"catAx" | b"valAx" | b"dateAx" | b"serAx"
                    ) =>
            {
                return Err(Error::Invalid(format!(
                    "chart plot-area element {} cannot be empty",
                    String::from_utf8_lossy(e.local_name().as_ref())
                )));
            },
            Ok(Event::Start(ref e)) => {
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
                    b"ofPieChart" => {
                        plot_area
                            .type_groups
                            .push(TypeGroup::OfPie(parse_of_pie_chart(reader)?));
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
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    Ok(plot_area)
}

fn is_chart_type_group_name(name: &[u8]) -> bool {
    matches!(
        name,
        b"areaChart"
            | b"area3DChart"
            | b"barChart"
            | b"bar3DChart"
            | b"bubbleChart"
            | b"doughnutChart"
            | b"lineChart"
            | b"line3DChart"
            | b"ofPieChart"
            | b"pieChart"
            | b"pie3DChart"
            | b"radarChart"
            | b"scatterChart"
            | b"stockChart"
            | b"surfaceChart"
            | b"surface3DChart"
    )
}

fn parse_data_table<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DataTable> {
    let mut data_table = DataTable::default();
    let mut seen = [false; 4];
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if data_table.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data table contains duplicate shape properties".into(),
                    ));
                }
                data_table.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart data-table shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if data_table.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data table contains duplicate shape properties".into(),
                    ));
                }
                data_table.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if data_table.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data table contains duplicate text properties".into(),
                    ));
                }
                data_table.text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart data-table text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if data_table.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data table contains duplicate text properties".into(),
                    ));
                }
                data_table.text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if data_table.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart data table contains duplicate extension lists".into(),
                    ));
                }
                data_table.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart data-table extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if data_table.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart data table contains duplicate extension lists".into(),
                    ));
                }
                data_table.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                let field = match element.local_name().as_ref() {
                    b"showHorzBorder" => Some((0, &mut data_table.show_horizontal_border)),
                    b"showVertBorder" => Some((1, &mut data_table.show_vertical_border)),
                    b"showOutline" => Some((2, &mut data_table.show_outline)),
                    b"showKeys" => Some((3, &mut data_table.show_legend_keys)),
                    _ => None,
                };
                if let Some((index, field)) = field {
                    if seen[index] {
                        return Err(Error::Invalid(
                            "chart data table contains a duplicate visibility setting".into(),
                        ));
                    }
                    seen[index] = true;
                    *field = parse_bool_attr(element)?;
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dTable" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart data table".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(data_table)
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
                return Err(Error::Invalid("unterminated chart layout".into()));
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
    supports_data_labels: bool,
    supports_axes: bool,
    mut drop_lines: Option<&mut Option<Lines>>,
    mut extra: impl FnMut(&BytesStart<'_>) -> Result<()>,
) -> Result<TypeGroupCommon> {
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element))
                if drop_lines.is_some() && element.local_name().as_ref() == b"dropLines" =>
            {
                let lines = parse_chart_lines(reader, b"dropLines")?;
                if let Some(target) = drop_lines.as_deref_mut() {
                    set_chart_lines(target, lines, "chart drop lines")?;
                }
            },
            Ok(Event::Empty(ref element))
                if drop_lines.is_some() && element.local_name().as_ref() == b"dropLines" =>
            {
                if let Some(target) = drop_lines.as_deref_mut() {
                    set_chart_lines(target, Lines::new(), "chart drop lines")?;
                }
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element))
                if supports_data_labels && element.local_name().as_ref() == b"dLbls" =>
            {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref element))
                if supports_data_labels && element.local_name().as_ref() == b"dLbls" =>
            {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"varyColors" => common.vary_colors = parse_bool_attr(element)?,
                    b"axId" if supports_axes => common
                        .axis_ids
                        .push(required_u32_attr(element, "chart type-group axis ID")?),
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
                return Err(Error::Invalid("unterminated chart type group".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(common)
}

fn parse_type_group_extension<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    common: &mut TypeGroupCommon,
    element: &BytesStart<'_>,
    empty: bool,
) -> Result<()> {
    if common.extension_list.is_some() {
        return Err(Error::Invalid(
            "chart type group contains duplicate extension lists".into(),
        ));
    }
    let xml = if empty {
        reader.capture_empty_fragment(element)?
    } else {
        reader.capture_fragment(element, "chart type-group extension list")?
    };
    common.extension_list = Some(ExtensionList::from_xml(xml)?);
    Ok(())
}

fn begin_group_data_labels(seen: &mut bool) -> Result<()> {
    if *seen {
        return Err(Error::Invalid(
            "chart type group contains duplicate data-label settings".into(),
        ));
    }
    *seen = true;
    Ok(())
}

fn set_chart_lines(target: &mut Option<Lines>, lines: Lines, description: &str) -> Result<()> {
    if target.replace(lines).is_some() {
        return Err(Error::Invalid(format!("{description} are duplicated")));
    }
    Ok(())
}

fn parse_chart_lines<R: BufRead>(reader: &mut ChartXmlReader<R>, end_name: &[u8]) -> Result<Lines> {
    let mut lines = Lines::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if lines.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart lines contain duplicate shape properties".into(),
                    ));
                }
                lines.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart-line shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if lines.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart lines contain duplicate shape properties".into(),
                    ));
                }
                lines.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == end_name => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart-line formatting".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(lines)
}

fn set_empty_up_down_bars(target: &mut Option<UpDownBars>, description: &str) -> Result<()> {
    if target.replace(UpDownBars::default()).is_some() {
        return Err(Error::Invalid(format!(
            "{description} contains duplicate up/down bars"
        )));
    }
    Ok(())
}

fn parse_area_3d_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Area3DTypeGroup> {
    let mut grouping = BarGrouping::Standard;
    let mut gap_depth = None;
    let mut drop_lines = None;
    let common = parse_common_type_group(
        reader,
        b"area3DChart",
        true,
        true,
        Some(&mut drop_lines),
        |element| {
            match element.local_name().as_ref() {
                b"grouping" => grouping = parse_grouping(element)?,
                b"gapDepth" => {
                    gap_depth = Some(match get_attr(element, b"val") {
                        Some(_) => {
                            bounded_percentage_u32_attr(element, "area 3D gap depth", 0, 500)?
                        },
                        None => 150,
                    });
                },
                _ => {},
            }
            Ok(())
        },
    )?;
    let mut group = Area3DTypeGroup::new(grouping);
    group.common = common;
    group.gap_depth = gap_depth;
    group.drop_lines = drop_lines;
    Ok(group)
}

fn parse_bubble_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<BubbleTypeGroup> {
    let mut group_bubble_3d = false;
    let mut scale = BubbleScale::default();
    let mut show_negative_bubbles = true;
    let mut size = BubbleSize::default();
    let common = parse_common_type_group(reader, b"bubbleChart", true, true, None, |element| {
        match element.local_name().as_ref() {
            b"bubble3D" => group_bubble_3d = parse_bool_attr(element)?,
            b"bubbleScale" => {
                let value = match get_attr(element, b"val") {
                    Some(_) => bounded_percentage_u32_attr(element, "bubble scale", 0, 300)?,
                    None => u32::from(BubbleScale::DEFAULT),
                };
                scale = BubbleScale::try_from(value)
                    .map_err(|error| Error::Invalid(error.to_string()))?;
            },
            b"showNegBubbles" => show_negative_bubbles = parse_bool_attr(element)?,
            b"sizeRepresents" => {
                let value = get_attr(element, b"val").unwrap_or_else(|| b"area".to_vec());
                size = BubbleSize::from_xml(&value)
                    .map_err(|_| invalid_attribute("chart bubble size representation", &value))?;
            },
            _ => {},
        }
        Ok(())
    })?;
    let mut group = BubbleTypeGroup::new();
    group.common = common;
    // ECMA-376 permits bubble3D at group level, but Microsoft Office rejects
    // that placement. Canonicalize its semantic effect onto each series,
    // which is the Office-compatible representation.
    if group_bubble_3d {
        for series in &mut group.common.series {
            series.bubble_3d = true;
        }
    }
    group.set_scale(scale);
    group.show_negative_bubbles = show_negative_bubbles;
    group.set_size(size);
    Ok(group)
}

fn parse_doughnut_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DoughnutTypeGroup> {
    let mut first_slice_angle = 0;
    let mut hole_size = 50;
    let common = parse_common_type_group(reader, b"doughnutChart", true, false, None, |element| {
        match element.local_name().as_ref() {
            b"firstSliceAng" => {
                first_slice_angle = match get_attr(element, b"val") {
                    Some(_) => required_u32_attr(element, "first-slice angle")?,
                    None => 0,
                };
                if first_slice_angle > 360 {
                    return Err(Error::Invalid(
                        "chart first-slice angle exceeds 360".to_string(),
                    ));
                }
            },
            b"holeSize" => {
                hole_size = match get_attr(element, b"val") {
                    Some(_) => bounded_percentage_u32_attr(element, "doughnut hole size", 1, 90)?,
                    None => 10,
                };
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
    let mut gap_depth = None;
    let mut drop_lines = None;
    let common = parse_common_type_group(
        reader,
        b"line3DChart",
        true,
        true,
        Some(&mut drop_lines),
        |element| {
            match element.local_name().as_ref() {
                b"grouping" => grouping = parse_grouping(element)?,
                b"gapDepth" => {
                    gap_depth = Some(match get_attr(element, b"val") {
                        Some(_) => {
                            bounded_percentage_u32_attr(element, "line 3D gap depth", 0, 500)?
                        },
                        None => 150,
                    });
                },
                _ => {},
            }
            Ok(())
        },
    )?;
    let mut group = Line3DTypeGroup::new(grouping);
    group.common = common;
    group.gap_depth = gap_depth;
    group.drop_lines = drop_lines;
    Ok(group)
}

fn parse_pie_3d_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Pie3DTypeGroup> {
    let mut group = Pie3DTypeGroup::new();
    group.common = parse_common_type_group(reader, b"pie3DChart", true, false, None, |_| Ok(()))?;
    Ok(group)
}

fn parse_of_pie_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<OfPieTypeGroup> {
    let mut group = OfPieTypeGroup::new(OfPieType::Pie);
    let mut saw_of_pie_type = false;
    let mut saw_gap_width = false;
    let mut saw_split_type = false;
    let mut saw_split_position = false;
    let mut saw_custom_split = false;
    let mut saw_second_pie_size = false;
    let mut saw_data_labels = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"serLines" => {
                group
                    .series_lines
                    .push(parse_chart_lines(reader, b"serLines")?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"serLines" => {
                group.series_lines.push(Lines::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut group.common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut group.common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                group.common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                group.common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"custSplit" => {
                if saw_custom_split {
                    return Err(Error::Invalid(
                        "of-pie chart contains duplicate custom splits".into(),
                    ));
                }
                saw_custom_split = true;
                group.custom_split_points = Some(parse_custom_pie_split(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"custSplit" => {
                if saw_custom_split {
                    return Err(Error::Invalid(
                        "of-pie chart contains duplicate custom splits".into(),
                    ));
                }
                saw_custom_split = true;
                group.custom_split_points = Some(Vec::new());
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"ofPieType" => {
                        if saw_of_pie_type {
                            return Err(Error::Invalid(
                                "of-pie chart contains duplicate plot types".into(),
                            ));
                        }
                        saw_of_pie_type = true;
                        group.of_pie_type = match get_attr(element, b"val").as_deref() {
                            None | Some(b"pie") => OfPieType::Pie,
                            Some(b"bar") => OfPieType::Bar,
                            Some(value) => return Err(invalid_attribute("of-pie type", value)),
                        };
                    },
                    b"varyColors" => group.common.vary_colors = parse_bool_attr(element)?,
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            group.common.series.push(series);
                        }
                    },
                    b"gapWidth" => {
                        if saw_gap_width {
                            return Err(Error::Invalid(
                                "of-pie chart contains duplicate gap widths".into(),
                            ));
                        }
                        saw_gap_width = true;
                        group.gap_width = Some(match get_attr(element, b"val") {
                            Some(_) => {
                                bounded_percentage_u32_attr(element, "of-pie gap width", 0, 500)?
                            },
                            None => 150,
                        });
                    },
                    b"splitType" => {
                        if saw_split_type {
                            return Err(Error::Invalid(
                                "of-pie chart contains duplicate split types".into(),
                            ));
                        }
                        saw_split_type = true;
                        group.split_type = Some(match get_attr(element, b"val").as_deref() {
                            None | Some(b"auto") => OfPieSplitType::Automatic,
                            Some(b"cust") => OfPieSplitType::Custom,
                            Some(b"percent") => OfPieSplitType::Percent,
                            Some(b"pos") => OfPieSplitType::Position,
                            Some(b"val") => OfPieSplitType::Value,
                            Some(value) => {
                                return Err(invalid_attribute("of-pie split type", value));
                            },
                        });
                    },
                    b"splitPos" => {
                        if saw_split_position {
                            return Err(Error::Invalid(
                                "of-pie chart contains duplicate split positions".into(),
                            ));
                        }
                        saw_split_position = true;
                        group.split_position =
                            Some(required_f64_attr(element, "of-pie split position")?);
                    },
                    b"secondPieSize" => {
                        if saw_second_pie_size {
                            return Err(Error::Invalid(
                                "of-pie chart contains duplicate secondary sizes".into(),
                            ));
                        }
                        saw_second_pie_size = true;
                        group.second_pie_size = Some(match get_attr(element, b"val") {
                            Some(_) => bounded_percentage_u32_attr(
                                element,
                                "of-pie secondary size",
                                5,
                                200,
                            )?,
                            None => 75,
                        });
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"ofPieChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated of-pie chart".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if !saw_of_pie_type {
        return Err(Error::Invalid(
            "of-pie chart is missing its plot type".into(),
        ));
    }
    Ok(group)
}

fn parse_custom_pie_split<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Vec<u32>> {
    let mut points = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"secondPiePt" =>
            {
                points.push(required_u32_attr(element, "of-pie secondary point")?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"custSplit" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated of-pie custom split".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(points)
}

fn parse_up_down_bars<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<UpDownBars> {
    let mut bars = UpDownBars::default();
    let mut saw_gap_width = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"upBars" => {
                let lines = parse_chart_lines(reader, b"upBars")?;
                set_chart_lines(&mut bars.up_bars, lines, "chart up bars")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"upBars" => {
                set_chart_lines(&mut bars.up_bars, Lines::new(), "chart up bars")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"downBars" => {
                let lines = parse_chart_lines(reader, b"downBars")?;
                set_chart_lines(&mut bars.down_bars, lines, "chart down bars")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"downBars" => {
                set_chart_lines(&mut bars.down_bars, Lines::new(), "chart down bars")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if bars.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart up/down bars contain duplicate extension lists".into(),
                    ));
                }
                bars.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart up/down-bar extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if bars.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart up/down bars contain duplicate extension lists".into(),
                    ));
                }
                bars.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"gapWidth" =>
            {
                if saw_gap_width {
                    return Err(Error::Invalid(
                        "chart up/down bars contain duplicate gap widths".into(),
                    ));
                }
                saw_gap_width = true;
                bars.gap_width = Some(match get_attr(element, b"val") {
                    Some(_) => {
                        bounded_percentage_u32_attr(element, "chart up/down-bar gap width", 0, 500)?
                    },
                    None => 150,
                });
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"upDownBars" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart up/down bars".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(bars)
}

fn parse_radar_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<RadarTypeGroup> {
    let mut style = RadarStyle::Standard;
    let common = parse_common_type_group(reader, b"radarChart", true, true, None, |element| {
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
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut drop_lines = None;
    let mut high_low_lines = None;
    let mut up_down_bars = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dropLines" => {
                let lines = parse_chart_lines(reader, b"dropLines")?;
                set_chart_lines(&mut drop_lines, lines, "stock chart drop lines")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dropLines" => {
                set_chart_lines(&mut drop_lines, Lines::new(), "stock chart drop lines")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"hiLowLines" => {
                let lines = parse_chart_lines(reader, b"hiLowLines")?;
                set_chart_lines(&mut high_low_lines, lines, "stock chart high/low lines")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"hiLowLines" => {
                set_chart_lines(
                    &mut high_low_lines,
                    Lines::new(),
                    "stock chart high/low lines",
                )?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"upDownBars" => {
                if up_down_bars.is_some() {
                    return Err(Error::Invalid(
                        "stock chart contains duplicate up/down bars".into(),
                    ));
                }
                up_down_bars = Some(parse_up_down_bars(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"upDownBars" => {
                set_empty_up_down_bars(&mut up_down_bars, "stock chart")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(element, "stock chart axis ID")?),
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"stockChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated stock chart".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = StockTypeGroup::new();
    group.common = common;
    group.drop_lines = drop_lines;
    group.high_low_lines = high_low_lines;
    group.up_down_bars = up_down_bars;
    Ok(group)
}

fn parse_surface_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<SurfaceTypeGroup> {
    let (common, wireframe, band_formats) = parse_surface_type_group(reader, b"surfaceChart")?;
    let mut group = SurfaceTypeGroup::new();
    group.common = common;
    group.wireframe = wireframe;
    group.band_formats = band_formats;
    Ok(group)
}

fn parse_surface_3d_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Surface3DTypeGroup> {
    let (common, wireframe, band_formats) = parse_surface_type_group(reader, b"surface3DChart")?;
    let mut group = Surface3DTypeGroup::new();
    group.common = common;
    group.wireframe = wireframe;
    group.band_formats = band_formats;
    Ok(group)
}

fn parse_surface_type_group<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    end_name: &[u8],
) -> Result<(TypeGroupCommon, bool, Option<Vec<BandFormat>>)> {
    let mut common = TypeGroupCommon::new();
    let mut wireframe = false;
    let mut band_formats = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"bandFmts" => {
                if band_formats.is_some() {
                    return Err(Error::Invalid(
                        "surface chart contains duplicate band-format collections".into(),
                    ));
                }
                band_formats = Some(parse_surface_band_formats(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"bandFmts" => {
                if band_formats.is_some() {
                    return Err(Error::Invalid(
                        "surface chart contains duplicate band-format collections".into(),
                    ));
                }
                band_formats = Some(Vec::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"ser" => {
                if let Some(series) = parse_series(reader)? {
                    common.series.push(series);
                }
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"ser" => {
                return Err(Error::Invalid(
                    "surface chart contains an empty series".into(),
                ));
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"wireframe" => wireframe = parse_bool_attr(element)?,
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(element, "surface chart axis ID")?),
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == end_name => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated surface chart".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok((common, wireframe, band_formats))
}

fn parse_surface_band_formats<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Vec<BandFormat>> {
    let mut formats = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"bandFmt" => {
                let format = parse_surface_band_format(reader)?;
                if formats
                    .iter()
                    .any(|existing: &BandFormat| existing.index == format.index)
                {
                    return Err(Error::Invalid(format!(
                        "surface chart contains duplicate band index {}",
                        format.index
                    )));
                }
                formats.push(format);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"bandFmt" => {
                return Err(Error::Invalid(
                    "surface chart band format is missing its index".into(),
                ));
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"bandFmts" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
                    "unterminated surface chart band formats".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(formats)
}

fn parse_surface_band_format<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<BandFormat> {
    let mut index = None;
    let mut shape_properties = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "surface chart band format contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "surface chart band shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "surface chart band format contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"idx" =>
            {
                if index.is_some() {
                    return Err(Error::Invalid(
                        "surface chart band format contains duplicate indexes".into(),
                    ));
                }
                index = Some(required_u32_attr(element, "surface chart band index")?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"bandFmt" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
                    "unterminated surface chart band format".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    let mut format =
        BandFormat::new(index.ok_or_else(|| missing_attribute("surface chart band index"))?);
    format.shape_properties = shape_properties;
    Ok(format)
}

fn parse_bar_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<BarTypeGroup>> {
    let mut direction = BarDirection::Column;
    let mut grouping = BarGrouping::Clustered;
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut gap_width = None;
    let mut overlap = None;
    let mut series_lines = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"serLines" => {
                series_lines.push(parse_chart_lines(reader, b"serLines")?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"serLines" => {
                series_lines.push(Lines::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
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
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(e, "bar chart axis ID")?),
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"gapWidth" => {
                        gap_width = Some(match get_attr(e, b"val") {
                            Some(_) => bounded_percentage_u32_attr(e, "chart gap width", 0, 500)?,
                            None => 150,
                        });
                    },
                    b"overlap" => {
                        overlap = Some(match get_attr(e, b"val") {
                            Some(_) => bounded_percentage_i32_attr(e, "chart overlap", -100, 100)?,
                            None => 0,
                        });
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"barChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
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
    group.series_lines = series_lines;
    Ok(Some(group))
}

fn parse_bar_3d_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<Bar3DTypeGroup>> {
    let mut direction = BarDirection::Column;
    let mut grouping = BarGrouping::Clustered;
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut gap_width = None;
    let mut gap_depth = None;
    let mut shape = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
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
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(e, "3D bar chart axis ID")?),
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"gapWidth" => {
                        gap_width = Some(match get_attr(e, b"val") {
                            Some(_) => bounded_percentage_u32_attr(e, "chart gap width", 0, 500)?,
                            None => 150,
                        });
                    },
                    b"gapDepth" => {
                        gap_depth = Some(match get_attr(e, b"val") {
                            Some(_) => bounded_percentage_u32_attr(e, "chart gap depth", 0, 500)?,
                            None => 150,
                        });
                    },
                    b"shape" => {
                        shape = Some(parse_bar_shape(e, "chart bar shape")?);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"bar3DChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
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
    let mut saw_data_labels = false;
    let mut marker = true;
    let mut smooth = false;
    let mut drop_lines = None;
    let mut high_low_lines = None;
    let mut up_down_bars = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dropLines" => {
                let lines = parse_chart_lines(reader, b"dropLines")?;
                set_chart_lines(&mut drop_lines, lines, "line chart drop lines")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dropLines" => {
                set_chart_lines(&mut drop_lines, Lines::new(), "line chart drop lines")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"hiLowLines" => {
                let lines = parse_chart_lines(reader, b"hiLowLines")?;
                set_chart_lines(&mut high_low_lines, lines, "line chart high/low lines")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"hiLowLines" => {
                set_chart_lines(
                    &mut high_low_lines,
                    Lines::new(),
                    "line chart high/low lines",
                )?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"upDownBars" => {
                if up_down_bars.is_some() {
                    return Err(Error::Invalid(
                        "line chart contains duplicate up/down bars".into(),
                    ));
                }
                up_down_bars = Some(parse_up_down_bars(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"upDownBars" => {
                set_empty_up_down_bars(&mut up_down_bars, "line chart")?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"grouping" => {
                        grouping = parse_grouping(e)?;
                    },
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(e, "line chart axis ID")?),
                    b"ser" => {
                        if let Some(series) = parse_series(reader)? {
                            common.series.push(series);
                        }
                    },
                    b"marker" => marker = parse_bool_attr(e)?,
                    b"smooth" => smooth = parse_bool_attr(e)?,
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"lineChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = LineTypeGroup::new(grouping);
    group.common = common;
    group.marker = marker;
    group.smooth = smooth;
    group.drop_lines = drop_lines;
    group.high_low_lines = high_low_lines;
    group.up_down_bars = up_down_bars;
    Ok(Some(group))
}

fn parse_pie_chart<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<PieTypeGroup>> {
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut first_slice_angle = 0;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
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
                        first_slice_angle = match get_attr(e, b"val") {
                            Some(_) => bounded_u32_attr(e, "chart first-slice angle", 0, 360)?,
                            None => 0,
                        };
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"pieChart" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
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
    let mut saw_data_labels = false;
    let mut drop_lines = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dropLines" => {
                let lines = parse_chart_lines(reader, b"dropLines")?;
                set_chart_lines(&mut drop_lines, lines, "area chart drop lines")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dropLines" => {
                set_chart_lines(&mut drop_lines, Lines::new(), "area chart drop lines")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"grouping" => {
                        grouping = parse_grouping(e)?;
                    },
                    b"varyColors" => {
                        common.vary_colors = parse_bool_attr(e)?;
                    },
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(e, "area chart axis ID")?),
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
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut group = AreaTypeGroup::new(grouping);
    group.common = common;
    group.drop_lines = drop_lines;
    Ok(Some(group))
}

fn parse_scatter_chart<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<Option<ScatterTypeGroup>> {
    let mut style = ScatterStyle::LineMarker;
    let mut common = TypeGroupCommon::new();
    let mut saw_data_labels = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_type_group_extension(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(parse_data_labels(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dLbls" => {
                begin_group_data_labels(&mut saw_data_labels)?;
                common.data_labels = Some(DataLabels::default());
            },
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
                    b"axId" => common
                        .axis_ids
                        .push(required_u32_attr(e, "scatter chart axis ID")?),
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
                return Err(Error::Invalid("unterminated chart element".to_string()));
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
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if series.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart series contains duplicate shape properties".into(),
                    ));
                }
                series.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart series shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if series.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart series contains duplicate shape properties".into(),
                    ));
                }
                series.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pictureOptions" => {
                if series.picture_options.is_some() {
                    return Err(Error::Invalid(
                        "chart series contains duplicate picture options".into(),
                    ));
                }
                series.picture_options = Some(parse_picture_options(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"pictureOptions" => {
                if series.picture_options.is_some() {
                    return Err(Error::Invalid(
                        "chart series contains duplicate picture options".into(),
                    ));
                }
                series.picture_options = Some(PictureOptions::default());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if series.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart series contains duplicate extension lists".into(),
                    ));
                }
                series.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(e, "chart series extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if series.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart series contains duplicate extension lists".into(),
                    ));
                }
                series.extension_list =
                    Some(ExtensionList::from_xml(reader.capture_empty_fragment(e)?)?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"marker" => {
                if saw_marker {
                    return Err(Error::Invalid(
                        "chart series has duplicate marker".to_string(),
                    ));
                }
                saw_marker = true;
                series.marker_present = true;
                let marker = parse_series_marker(reader)?;
                series.marker_symbol = marker.symbol;
                series.marker_size = marker.size;
                series.marker_shape_properties = marker.shape_properties;
                series.marker_extension_list = marker.extension_list;
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"marker" => {
                if saw_marker {
                    return Err(Error::Invalid(
                        "chart series has duplicate marker".to_string(),
                    ));
                }
                saw_marker = true;
                series.marker_present = true;
            },
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag_name = e.local_name();
                match tag_name.as_ref() {
                    b"idx" => {
                        if saw_index {
                            return Err(Error::Invalid(
                                "chart series has duplicate index".to_string(),
                            ));
                        }
                        saw_index = true;
                        series.index = required_u32_attr(e, "chart series index")?;
                    },
                    b"order" => {
                        if saw_order {
                            return Err(Error::Invalid(
                                "chart series has duplicate order".to_string(),
                            ));
                        }
                        saw_order = true;
                        series.order = required_u32_attr(e, "chart series order")?;
                    },
                    b"tx" => {
                        series.title = parse_series_title(reader)?;
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
                            return Err(Error::Invalid(format!(
                                "chart series has duplicate data-point index {}",
                                point.index
                            )));
                        }
                        series.data_points.push(point);
                    },
                    b"dLbls" => {
                        if saw_data_labels {
                            return Err(Error::Invalid(
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
                            return Err(Error::Invalid(
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
                    b"shape" => {
                        if series.bar_shape.is_some() {
                            return Err(Error::Invalid(
                                "chart series contains duplicate bar shapes".into(),
                            ));
                        }
                        series.bar_shape = Some(parse_bar_shape(e, "chart series bar shape")?);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"ser" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if !saw_index || !saw_order {
        return Err(Error::Invalid(
            "chart series requires both index and order".to_string(),
        ));
    }
    Ok(Some(series))
}

fn parse_bar_shape(element: &BytesStart<'_>, description: &str) -> Result<BarShape> {
    let value = get_attr(element, b"val").unwrap_or_else(|| b"box".to_vec());
    match value.as_slice() {
        b"box" => Ok(BarShape::Box),
        b"cone" => Ok(BarShape::Cone),
        b"coneToMax" => Ok(BarShape::ConeToMax),
        b"cylinder" => Ok(BarShape::Cylinder),
        b"pyramid" => Ok(BarShape::Pyramid),
        b"pyramidToMax" => Ok(BarShape::PyramidToMax),
        _ => Err(invalid_attribute(description, &value)),
    }
}

fn parse_series_marker<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Marker> {
    let mut symbol = None;
    let mut size = None;
    let mut shape_properties = None;
    let mut extension_list = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart marker has duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart marker shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart marker has duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart marker has duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart marker extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart marker has duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"symbol" => {
                        if symbol.is_some() {
                            return Err(Error::Invalid(
                                "chart marker has duplicate symbol".to_string(),
                            ));
                        }
                        symbol = Some(parse_marker_style(element)?);
                    },
                    b"size" => {
                        if size.is_some() {
                            return Err(Error::Invalid(
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
                return Err(Error::Invalid(
                    "unterminated chart series marker".to_string(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(Marker {
        symbol,
        size,
        shape_properties,
        extension_list,
    })
}

fn parse_data_point<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DataPoint> {
    let mut index = None;
    let mut explosion = None;
    let mut marker = None;
    let mut invert_if_negative = false;
    let mut bubble_3d = None;
    let mut shape_properties = None;
    let mut picture_options = None;
    let mut extension_list = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"marker" => {
                if marker.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate markers".into(),
                    ));
                }
                marker = Some(parse_series_marker(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"marker" => {
                if marker.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate markers".into(),
                    ));
                }
                marker = Some(Marker::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart data-point shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"pictureOptions" => {
                if picture_options.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate picture options".into(),
                    ));
                }
                picture_options = Some(parse_picture_options(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"pictureOptions" => {
                if picture_options.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate picture options".into(),
                    ));
                }
                picture_options = Some(PictureOptions::default());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart data-point extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart data point contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
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
                return Err(Error::Invalid("unterminated chart data point".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    let mut point =
        DataPoint::new(index.ok_or_else(|| missing_attribute("chart data-point index"))?);
    point.explosion = explosion;
    if let Some(marker) = marker {
        point.marker_present = true;
        point.marker_size = marker.size;
        point.marker_symbol = marker.symbol;
        point.marker_shape_properties = marker.shape_properties;
        point.marker_extension_list = marker.extension_list;
    }
    point.invert_if_negative = invert_if_negative;
    point.bubble_3d = bubble_3d;
    point.shape_properties = shape_properties;
    point.picture_options = picture_options;
    point.extension_list = extension_list;
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
                    return Err(Error::Invalid(format!(
                        "chart data labels contain duplicate point index {}",
                        label.index
                    )));
                }
                labels.labels.push(label);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dLbl" => {
                return Err(Error::Invalid(
                    "chart point data label is missing its index".into(),
                ));
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if labels.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data labels contain duplicate shape properties".into(),
                    ));
                }
                labels.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart data-label shape properties")?,
                )?);
                saw_shared_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if labels.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data labels contain duplicate shape properties".into(),
                    ));
                }
                labels.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
                saw_shared_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if labels.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data labels contain duplicate text properties".into(),
                    ));
                }
                labels.text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart data-label text properties")?,
                )?);
                saw_shared_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if labels.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart data labels contain duplicate text properties".into(),
                    ));
                }
                labels.text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
                saw_shared_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"leaderLines" => {
                set_chart_lines(
                    &mut labels.leader_lines,
                    parse_chart_lines(reader, b"leaderLines")?,
                    "chart data-label leader lines",
                )?;
                saw_shared_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"leaderLines" => {
                set_chart_lines(
                    &mut labels.leader_lines,
                    Lines::new(),
                    "chart data-label leader lines",
                )?;
                saw_shared_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if labels.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart data labels contain duplicate extension lists".into(),
                    ));
                }
                labels.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart data-label extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if labels.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart data labels contain duplicate extension lists".into(),
                    ));
                }
                labels.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"separator" => {
                saw_shared_settings = true;
                labels.separator = Some(parse_text_element(reader, b"separator")?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"delete" => {
                        if saw_delete {
                            return Err(Error::Invalid(
                                "chart data labels contain duplicate delete flags".into(),
                            ));
                        }
                        labels.deleted = parse_bool_attr(element)?;
                        saw_delete = true;
                    },
                    b"numFmt" => {
                        if saw_number_format {
                            return Err(Error::Invalid(
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
                return Err(Error::Invalid("unterminated chart data labels".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    if saw_delete && saw_shared_settings {
        return Err(Error::Invalid(
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
                    return Err(Error::Invalid(
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
                        return Err(Error::Invalid(
                            "chart point data label contains duplicate layouts".into(),
                        ));
                    },
                });
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"tx" => {
                if label.text.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate text".into(),
                    ));
                }
                label.text = parse_label_text(reader)?;
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if label.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate shape properties".into(),
                    ));
                }
                label.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart point data-label shape properties")?,
                )?);
                saw_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if label.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate shape properties".into(),
                    ));
                }
                label.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if label.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate text properties".into(),
                    ));
                }
                label.text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart point data-label text properties")?,
                )?);
                saw_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if label.text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate text properties".into(),
                    ));
                }
                label.text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if label.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate extension lists".into(),
                    ));
                }
                label.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart point data-label extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if label.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart point data label contains duplicate extension lists".into(),
                    ));
                }
                label.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"separator" => {
                label.separator = Some(parse_text_element(reader, b"separator")?);
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                let is_setting = match element.local_name().as_ref() {
                    b"idx" => {
                        if saw_index {
                            return Err(Error::Invalid(
                                "chart point data label contains duplicate indexes".into(),
                            ));
                        }
                        label.index = required_u32_attr(element, "chart point data-label index")?;
                        saw_index = true;
                        false
                    },
                    b"delete" => {
                        if saw_delete {
                            return Err(Error::Invalid(
                                "chart point data label contains duplicate delete flags".into(),
                            ));
                        }
                        label.deleted = parse_bool_attr(element)?;
                        saw_delete = true;
                        false
                    },
                    b"numFmt" => {
                        if label.number_format.is_some() {
                            return Err(Error::Invalid(
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
                return Err(Error::Invalid("unterminated chart point data label".into()));
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
        return Err(Error::Invalid(
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
                    return Err(Error::Invalid(
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
                    .map_err(|error| Error::Xml(error.to_string()))?,
            ),
            Ok(Event::CData(value)) if in_text => text.push_str(
                &value
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?,
            ),
            Ok(Event::GeneralRef(reference)) if in_text => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"t" => {
                in_text = false;
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"tx" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart data-label text".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if let Some(formula) = formula {
        if saw_text {
            return Err(Error::Invalid(
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
            Ok(Event::Start(ref e))
                if matches!(e.local_name().as_ref(), b"name" | b"trendlineName") =>
            {
                if trendline.name.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate names".into(),
                    ));
                }
                trendline.name = Some(parse_text_element(reader, e.local_name().as_ref())?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if trendline.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate shape properties".into(),
                    ));
                }
                trendline.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart trendline shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if trendline.shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate shape properties".into(),
                    ));
                }
                trendline.shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if trendline.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate extension lists".into(),
                    ));
                }
                trendline.extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(e, "chart trendline extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if trendline.extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate extension lists".into(),
                    ));
                }
                trendline.extension_list =
                    Some(ExtensionList::from_xml(reader.capture_empty_fragment(e)?)?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"trendlineLbl" => {
                if trendline.show_label {
                    return Err(Error::Invalid(
                        "chart trendline contains duplicate labels".into(),
                    ));
                }
                trendline.show_label = true;
                let label = parse_trendline_label(reader)?;
                trendline.label = label.text;
                trendline.label_layout = label.layout;
                trendline.label_number_format = label.number_format;
                trendline.label_shape_properties = label.shape_properties;
                trendline.label_text_properties = label.text_properties;
                trendline.label_extension_list = label.extension_list;
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"trendlineLbl" => {
                if trendline.show_label {
                    return Err(Error::Invalid(
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
                return Err(Error::Invalid("unterminated chart trendline".into()));
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
        return Err(Error::Invalid(
            "only polynomial trendlines can specify an order".to_string(),
        ));
    }
    if !matches!(trendline.trendline_type, TrendlineType::MovingAverage)
        && trendline.period.is_some()
    {
        return Err(Error::Invalid(
            "only moving-average trendlines can specify a period".to_string(),
        ));
    }
    Ok(trendline)
}

#[derive(Default)]
struct ParsedTrendlineLabel {
    text: Option<TitleText>,
    layout: Option<Layout>,
    number_format: Option<NumberFormat>,
    shape_properties: Option<ShapeProperties>,
    text_properties: Option<TextProperties>,
    extension_list: Option<ExtensionList>,
}

fn parse_trendline_label<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<ParsedTrendlineLabel> {
    let mut text = None;
    let mut saw_text = false;
    let mut layout = None;
    let mut number_format = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut extension_list = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if layout.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate layouts".into(),
                    ));
                }
                layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                layout = Some(match layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(Error::Invalid(
                            "chart trendline label contains duplicate layouts".into(),
                        ));
                    },
                });
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"tx" => {
                if saw_text {
                    return Err(Error::Invalid(
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
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate number formats".into(),
                    ));
                }
                number_format = Some(parse_number_format(
                    element,
                    reader.decoder(),
                    "chart trendline-label",
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart trendline-label shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart trendline-label text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart trendline-label extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart trendline label contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"trendlineLbl" => {
                break;
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart trendline label".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(ParsedTrendlineLabel {
        text,
        layout,
        number_format,
        shape_properties,
        text_properties,
        extension_list,
    })
}

fn parse_error_bar<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<ErrorBar> {
    let mut direction = None;
    let mut error_type = None;
    let mut value_type = None;
    let mut value = None;
    let mut plus_values = None;
    let mut minus_values = None;
    let mut no_end_cap = false;
    let mut shape_properties = None;
    let mut extension_list = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart error bars contain duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart error-bar shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart error bars contain duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart error bars contain duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(e, "chart error-bar extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart error bars contain duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(reader.capture_empty_fragment(e)?)?);
            },
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
                return Err(Error::Invalid("unterminated chart error bars".into()));
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
        shape_properties,
        extension_list,
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
            return Err(Error::Invalid(
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
                return Err(Error::Invalid("unterminated chart element".to_string()));
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
                return Err(Error::Invalid("unterminated chart element".to_string()));
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
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"strRef" => {
                let formula = parse_series_title_reference(reader)?;
                set_title(
                    &mut title,
                    TitleText::Reference(DataSourceRef::new(formula)),
                )?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"rich" => {
                let text = parse_series_title_rich_text(reader)?;
                set_title(&mut title, TitleText::Literal(RichText::new(text)))?;
            },
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
                return Err(Error::Invalid(
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

/// Read a series title reference and ignore its optional cached string data.
///
/// Spreadsheet applications commonly emit both `<c:f>` and a `<c:strCache>`
/// under `<c:strRef>`. The cache is a derived value, not a second title source;
/// the typed model retains the authoritative reference and avoids rejecting a
/// valid Office-produced chart as a duplicate title.
fn parse_series_title_reference<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<String> {
    let mut formula = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"f" => {
                if formula.is_some() {
                    return Err(Error::Invalid(
                        "chart series title reference contains duplicate formulas".into(),
                    ));
                }
                formula = Some(parse_text_element(reader, b"f")?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"strRef" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
                    "unterminated chart series title reference".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    formula.ok_or_else(|| Error::Invalid("chart series title reference requires a formula".into()))
}

fn parse_series_title_rich_text<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<String> {
    let mut text = String::new();
    let mut in_text = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"t" => {
                in_text = true;
            },
            Ok(Event::Text(value)) if in_text => {
                text.push_str(
                    &value
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Ok(Event::CData(value)) if in_text => {
                text.push_str(
                    &value
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Ok(Event::GeneralRef(reference)) if in_text => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"t" => {
                in_text = false;
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"rich" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart series rich text".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(text)
}

fn set_title(target: &mut Option<TitleText>, title: TitleText) -> Result<()> {
    if target.replace(title).is_some() {
        return Err(Error::Invalid(
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
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Ok(Event::CData(value)) => {
                text.push_str(
                    &value
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Ok(Event::GeneralRef(reference)) => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == end_name => break,
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == IGNORED_NAMESPACE_ELEMENT.as_bytes() => {},
            Ok(Event::Start(_)) | Ok(Event::Empty(_)) => {
                return Err(Error::Invalid(
                    "chart text element contains nested markup".to_string(),
                ));
            },
            Ok(Event::Eof) => {
                return Err(Error::Invalid(
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
                text.push_str(&e.decode().map_err(|e| Error::Xml(e.to_string()))?);
            },
            Ok(Event::CData(e)) if in_v => {
                text.push_str(&e.decode().map_err(|e| Error::Xml(e.to_string()))?);
            },
            Ok(Event::GeneralRef(reference)) if in_v => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"v" => {
                in_v = false;
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"pt" => break,
            Ok(Event::Eof) => {
                return Err(Error::Invalid("unterminated chart element".to_string()));
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
            Error::Invalid(format!("invalid chart numeric point '{text}'"))
        })?))
    } else {
        Ok(None)
    }
}

#[derive(Default)]
struct ParsedAxisScaling {
    min: Option<f64>,
    max: Option<f64>,
    log_base: Option<f64>,
}

impl ParsedAxisScaling {
    fn parse_element(&mut self, element: &BytesStart<'_>) -> Result<bool> {
        let (slot, description) = match element.local_name().as_ref() {
            b"min" => (&mut self.min, "chart axis minimum"),
            b"max" => (&mut self.max, "chart axis maximum"),
            b"logBase" => (&mut self.log_base, "chart logarithmic base"),
            _ => return Ok(false),
        };
        if slot.is_some() {
            return Err(Error::Invalid(format!(
                "{description} is specified more than once"
            )));
        }
        let value = required_f64_attr(element, description)?;
        if element.local_name().as_ref() == b"logBase" && !(2.0..=1000.0).contains(&value) {
            return Err(Error::Invalid(
                "chart logarithmic base must be between 2 and 1000".into(),
            ));
        }
        *slot = Some(value);
        Ok(true)
    }

    fn validate(&self) -> Result<()> {
        if self.min.zip(self.max).is_some_and(|(min, max)| min > max) {
            return Err(Error::Invalid("chart axis minimum exceeds maximum".into()));
        }
        Ok(())
    }
}

struct ParsedAxisCommon {
    axis_id: Option<u32>,
    cross_axis_id: Option<u32>,
    position: Option<AxisPosition>,
    title: Option<TitleText>,
    title_layout: Option<Layout>,
    title_overlay: bool,
    title_shape_properties: Option<ShapeProperties>,
    title_text_properties: Option<TextProperties>,
    title_extension_list: Option<ExtensionList>,
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
    major_gridlines: Option<Lines>,
    minor_gridlines: Option<Lines>,
    shape_properties: Option<ShapeProperties>,
    text_properties: Option<TextProperties>,
    scaling_extension_list: Option<ExtensionList>,
    extension_list: Option<ExtensionList>,
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
            title_shape_properties: None,
            title_text_properties: None,
            title_extension_list: None,
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
            major_gridlines: None,
            minor_gridlines: None,
            shape_properties: None,
            text_properties: None,
            scaling_extension_list: None,
            extension_list: None,
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
        common.title_shape_properties = self.title_shape_properties;
        common.title_text_properties = self.title_text_properties;
        common.title_extension_list = self.title_extension_list;
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
        common.major_gridlines = self.major_gridlines;
        common.minor_gridlines = self.minor_gridlines;
        common.shape_properties = self.shape_properties;
        common.text_properties = self.text_properties;
        common.scaling_extension_list = self.scaling_extension_list;
        common.extension_list = self.extension_list;
        Ok(common)
    }
}

fn parse_axis_title<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    common: &mut ParsedAxisCommon,
) -> Result<()> {
    if common.title.is_some() {
        return Err(Error::Invalid(
            "chart axis contains duplicate titles".into(),
        ));
    }
    let title = parse_title(reader)?;
    common.title = Some(title.text);
    common.title_layout = title.layout;
    common.title_overlay = title.overlay;
    common.title_shape_properties = title.shape_properties;
    common.title_text_properties = title.text_properties;
    common.title_extension_list = title.extension_list;
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

fn is_axis_common_fragment(element: &BytesStart<'_>) -> bool {
    matches!(
        element.local_name().as_ref(),
        b"majorGridlines" | b"minorGridlines" | b"spPr" | b"txPr"
    )
}

fn parse_axis_common_fragment<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    common: &mut ParsedAxisCommon,
    element: &BytesStart<'_>,
    empty: bool,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"majorGridlines" => {
            let lines = if empty {
                Lines::new()
            } else {
                parse_chart_lines(reader, b"majorGridlines")?
            };
            set_chart_lines(
                &mut common.major_gridlines,
                lines,
                "chart axis major gridlines",
            )?;
            common.show_major_gridlines = true;
        },
        b"minorGridlines" => {
            let lines = if empty {
                Lines::new()
            } else {
                parse_chart_lines(reader, b"minorGridlines")?
            };
            set_chart_lines(
                &mut common.minor_gridlines,
                lines,
                "chart axis minor gridlines",
            )?;
            common.show_minor_gridlines = true;
        },
        b"spPr" => {
            if common.shape_properties.is_some() {
                return Err(Error::Invalid(
                    "chart axis contains duplicate shape properties".into(),
                ));
            }
            let xml = if empty {
                reader.capture_empty_fragment(element)?
            } else {
                reader.capture_fragment(element, "chart axis shape properties")?
            };
            common.shape_properties = Some(ShapeProperties::from_xml(xml)?);
        },
        b"txPr" => {
            if common.text_properties.is_some() {
                return Err(Error::Invalid(
                    "chart axis contains duplicate text properties".into(),
                ));
            }
            let xml = if empty {
                reader.capture_empty_fragment(element)?
            } else {
                reader.capture_fragment(element, "chart axis text properties")?
            };
            common.text_properties = Some(TextProperties::from_xml(xml)?);
        },
        _ => {},
    }
    Ok(())
}

fn parse_axis_extension<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    common: &mut ParsedAxisCommon,
    element: &BytesStart<'_>,
    empty: bool,
    scaling: bool,
) -> Result<()> {
    let target = if scaling {
        &mut common.scaling_extension_list
    } else {
        &mut common.extension_list
    };
    if target.is_some() {
        return Err(Error::Invalid(format!(
            "chart axis contains duplicate {} extension lists",
            if scaling { "scaling" } else { "axis" }
        )));
    }
    let xml = if empty {
        reader.capture_empty_fragment(element)?
    } else {
        reader.capture_fragment(
            element,
            if scaling {
                "chart axis scaling extension list"
            } else {
                "chart axis extension list"
            },
        )?
    };
    *target = Some(ExtensionList::from_xml(xml)?);
    Ok(())
}

fn parse_category_axis<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<CategoryAxis>> {
    let mut common = ParsedAxisCommon::new();
    let mut scaling = ParsedAxisScaling::default();
    let mut auto = true;
    let mut label_align = None;
    let mut label_offset = None;
    let mut tick_label_skip = None;
    let mut tick_mark_skip = None;
    let mut no_multi_level = false;
    let mut in_scaling = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, false, in_scaling)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, true, in_scaling)?;
            },
            Ok(Event::Start(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"title" => {
                parse_axis_title(reader, &mut common)?;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if !parse_axis_common_element(&mut common, element, reader.decoder())? =>
            {
                if in_scaling && scaling.parse_element(element)? {
                    buf.clear();
                    continue;
                }
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
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = false;
            },
            Ok(Event::Eof) => return Err(unterminated_axis("category")),
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    scaling.validate()?;
    let common = common.finish()?;
    let mut axis = CategoryAxis::new(common.axis_id, common.position, common.cross_axis_id);
    axis.common = common;
    axis.min = scaling.min;
    axis.max = scaling.max;
    axis.log_base = scaling.log_base;
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
    let mut scaling = ParsedAxisScaling::default();
    let mut major_unit = None;
    let mut minor_unit = None;
    let mut display_units = None;
    let mut cross_between = AxisCrossBetween::Between;
    let mut in_scaling = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, false, in_scaling)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, true, in_scaling)?;
            },
            Ok(Event::Start(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"title" => {
                parse_axis_title(reader, &mut common)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dispUnits" => {
                if display_units.is_some() {
                    return Err(Error::Invalid(
                        "chart value axis contains duplicate display units".into(),
                    ));
                }
                display_units = Some(parse_display_units(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dispUnits" => {
                return Err(Error::Invalid(
                    "chart display units are missing their unit".into(),
                ));
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if !parse_axis_common_element(&mut common, element, reader.decoder())? =>
            {
                if in_scaling && scaling.parse_element(element)? {
                    buf.clear();
                    continue;
                }
                match element.local_name().as_ref() {
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
                    b"crossBetween" => cross_between = parse_axis_cross_between(element)?,
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"valAx" => break,
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = false;
            },
            Ok(Event::Eof) => return Err(unterminated_axis("value")),
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    scaling.validate()?;
    let common = common.finish()?;
    let mut axis = ValueAxis::new(common.axis_id, common.position, common.cross_axis_id);
    axis.common = common;
    axis.min = scaling.min;
    axis.max = scaling.max;
    axis.major_unit = major_unit;
    axis.minor_unit = minor_unit;
    axis.log_base = scaling.log_base;
    axis.cross_between = cross_between;
    axis.display_units = display_units.map(Box::new);
    Ok(Some(axis))
}

fn parse_display_units<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<DisplayUnits> {
    let mut built_in_unit = None;
    let mut custom_unit = None;
    let mut parsed_label = ParsedDisplayUnitsLabel::default();
    let mut extension_list = None;
    let mut saw_label = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dispUnitsLbl" => {
                if saw_label {
                    return Err(Error::Invalid(
                        "chart display units contain duplicate labels".into(),
                    ));
                }
                saw_label = true;
                parsed_label = parse_display_units_label(reader)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dispUnitsLbl" => {
                if saw_label {
                    return Err(Error::Invalid(
                        "chart display units contain duplicate labels".into(),
                    ));
                }
                saw_label = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart display units contain duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart display-units extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart display units contain duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"builtInUnit" => {
                        if built_in_unit.is_some() {
                            return Err(Error::Invalid(
                                "chart display units contain duplicate built-in units".into(),
                            ));
                        }
                        built_in_unit = Some(parse_built_in_unit(element)?);
                    },
                    b"custUnit" => {
                        if custom_unit.is_some() {
                            return Err(Error::Invalid(
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
                return Err(Error::Invalid("unterminated chart display units".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if built_in_unit.is_some() == custom_unit.is_some() {
        return Err(Error::Invalid(
            "chart display units require exactly one built-in or custom unit".into(),
        ));
    }
    Ok(DisplayUnits {
        built_in_unit,
        custom_unit,
        show_label: saw_label,
        label: parsed_label.label,
        layout: parsed_label.layout,
        label_shape_properties: parsed_label.shape_properties,
        label_text_properties: parsed_label.text_properties,
        extension_list,
    })
}

#[derive(Default)]
struct ParsedDisplayUnitsLabel {
    label: Option<TitleText>,
    layout: Option<Layout>,
    shape_properties: Option<ShapeProperties>,
    text_properties: Option<TextProperties>,
}

fn parse_display_units_label<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
) -> Result<ParsedDisplayUnitsLabel> {
    let mut text = String::new();
    let mut formula = None;
    let mut layout = None;
    let mut in_text = false;
    let mut saw_text = false;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if layout.is_some() {
                    return Err(Error::Invalid(
                        "chart display-units label contains duplicate layouts".into(),
                    ));
                }
                layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                layout = Some(match layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(Error::Invalid(
                            "chart display-units label contains duplicate layouts".into(),
                        ));
                    },
                });
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart display-units label contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader
                        .capture_fragment(element, "chart display-units label shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart display-units label contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart display-units label contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader
                        .capture_fragment(element, "chart display-units label text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart display-units label contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"f" => {
                if formula.is_some() {
                    return Err(Error::Invalid(
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
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Ok(Event::CData(value)) if in_text => {
                text.push_str(
                    &value
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?,
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
                return Err(Error::Invalid(
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
            return Err(Error::Invalid(
                "chart display-units label mixes a formula reference with literal text".into(),
            ));
        }
        Some(TitleText::Reference(DataSourceRef::new(formula)))
    } else if saw_text {
        Some(TitleText::Literal(RichText::new(text)))
    } else {
        None
    };
    Ok(ParsedDisplayUnitsLabel {
        label,
        layout,
        shape_properties,
        text_properties,
    })
}

fn parse_date_axis<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<DateAxis>> {
    let mut common = ParsedAxisCommon::new();
    let mut scaling = ParsedAxisScaling::default();
    let mut major_unit = None;
    let mut minor_unit = None;
    let mut major_time_unit = None;
    let mut minor_time_unit = None;
    let mut base_time_unit = None;
    let mut auto = true;
    let mut label_offset = None;
    let mut in_scaling = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, false, in_scaling)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, true, in_scaling)?;
            },
            Ok(Event::Start(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"title" => {
                parse_axis_title(reader, &mut common)?;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if !parse_axis_common_element(&mut common, element, reader.decoder())? =>
            {
                if in_scaling && scaling.parse_element(element)? {
                    buf.clear();
                    continue;
                }
                match element.local_name().as_ref() {
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
                    b"lblOffset" => {
                        label_offset = Some(bounded_u32_attr(
                            element,
                            "chart date-axis label offset",
                            0,
                            1000,
                        )?);
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dateAx" => break,
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = false;
            },
            Ok(Event::Eof) => return Err(unterminated_axis("date")),
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    scaling.validate()?;
    let common = common.finish()?;
    let mut axis = DateAxis::new(common.axis_id, common.position, common.cross_axis_id);
    axis.common = common;
    axis.min = scaling.min;
    axis.max = scaling.max;
    axis.log_base = scaling.log_base;
    axis.major_unit = major_unit;
    axis.minor_unit = minor_unit;
    axis.major_time_unit = major_time_unit;
    axis.minor_time_unit = minor_time_unit;
    axis.base_time_unit = base_time_unit;
    axis.auto = auto;
    axis.label_offset = label_offset;
    Ok(Some(axis))
}

fn parse_series_axis<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Option<SeriesAxis>> {
    let mut common = ParsedAxisCommon::new();
    let mut scaling = ParsedAxisScaling::default();
    let mut tick_label_skip = None;
    let mut tick_mark_skip = None;
    let mut in_scaling = false;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, false, in_scaling)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                parse_axis_extension(reader, &mut common, element, true, in_scaling)?;
            },
            Ok(Event::Start(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, false)?;
            },
            Ok(Event::Empty(ref element)) if is_axis_common_fragment(element) => {
                parse_axis_common_fragment(reader, &mut common, element, true)?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"title" => {
                parse_axis_title(reader, &mut common)?;
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if !parse_axis_common_element(&mut common, element, reader.decoder())? =>
            {
                if in_scaling && scaling.parse_element(element)? {
                    buf.clear();
                    continue;
                }
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
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"scaling" => {
                in_scaling = false;
            },
            Ok(Event::Eof) => return Err(unterminated_axis("series")),
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    scaling.validate()?;
    let common = common.finish()?;
    let mut axis = SeriesAxis::new(common.axis_id, common.position, common.cross_axis_id);
    axis.common = common;
    axis.min = scaling.min;
    axis.max = scaling.max;
    axis.log_base = scaling.log_base;
    axis.tick_label_skip = tick_label_skip;
    axis.tick_mark_skip = tick_mark_skip;
    Ok(Some(axis))
}

fn unterminated_axis(kind: &str) -> Error {
    Error::Invalid(format!("unterminated chart {kind} axis"))
}

fn parse_legend<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<Legend> {
    let mut position = LegendPosition::Right;
    let mut overlay = false;
    let mut layout = None;
    let mut entries = Vec::new();
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut extension_list = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"legendEntry" => {
                let entry = parse_legend_entry(reader)?;
                if entries
                    .iter()
                    .any(|existing: &LegendEntry| existing.index == entry.index)
                {
                    return Err(Error::Invalid(format!(
                        "chart legend contains duplicate entry index {}",
                        entry.index
                    )));
                }
                entries.push(entry);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"legendEntry" => {
                return Err(Error::Invalid(
                    "chart legend entry is missing its index".into(),
                ));
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"layout" => {
                if layout.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate layouts".into(),
                    ));
                }
                layout = Some(parse_layout(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"layout" => {
                layout = Some(match layout {
                    None => Layout::new(),
                    Some(_) => {
                        return Err(Error::Invalid(
                            "chart legend contains duplicate layouts".into(),
                        ));
                    },
                });
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart legend shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart legend text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart legend extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart legend contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
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
                return Err(Error::Invalid("unterminated chart element".to_string()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let mut legend = Legend::new(position).with_overlay(overlay);
    legend.layout = layout;
    legend.entries = entries;
    legend.shape_properties = shape_properties;
    legend.text_properties = text_properties;
    legend.extension_list = extension_list;
    Ok(legend)
}

fn parse_legend_entry<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<LegendEntry> {
    let mut index = None;
    let mut deleted = false;
    let mut saw_delete = false;
    let mut text_properties = None;
    let mut extension_list = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if saw_delete || text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart legend entry contains multiple choice values".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_fragment(element, "chart legend-entry text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if saw_delete || text_properties.is_some() {
                    return Err(Error::Invalid(
                        "chart legend entry contains multiple choice values".into(),
                    ));
                }
                text_properties = Some(TextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart legend entry contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_fragment(element, "chart legend-entry extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(Error::Invalid(
                        "chart legend entry contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"idx" => {
                        if index.is_some() {
                            return Err(Error::Invalid(
                                "chart legend entry contains duplicate indexes".into(),
                            ));
                        }
                        index = Some(required_u32_attr(element, "chart legend entry index")?);
                    },
                    b"delete" => {
                        if saw_delete || text_properties.is_some() {
                            return Err(Error::Invalid(
                                "chart legend entry contains multiple choice values".into(),
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
                return Err(Error::Invalid("unterminated chart legend entry".into()));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    let index = index.ok_or_else(|| missing_attribute("chart legend entry index"))?;
    if !saw_delete && text_properties.is_none() {
        return Err(Error::Invalid(
            "chart legend entry is missing its delete or text-properties choice".into(),
        ));
    }
    let mut entry = LegendEntry::new(index);
    entry.deleted = deleted;
    entry.text_properties = text_properties;
    entry.extension_list = extension_list;
    Ok(entry)
}

#[inline]
fn parse_grouping(e: &BytesStart<'_>) -> Result<BarGrouping> {
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
fn parse_axis_position(e: &BytesStart<'_>) -> Result<AxisPosition> {
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
        .map_err(|error| Error::Xml(error.to_string()))?
        .ok_or_else(|| missing_attribute(&format!("{description} number format code")))?
        .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
        .map_err(|error| Error::Xml(error.to_string()))?
        .into_owned();
    let source_linked = match get_attr(element, b"sourceLinked") {
        Some(value) => parse_bool_value(&value, &format!("{description} source-linked flag"))?,
        None => true,
    };
    Ok(NumberFormat::new(format_code).with_source_linked(source_linked))
}

#[inline]
fn parse_display_blanks(e: &BytesStart<'_>) -> Result<DisplayBlanks> {
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
fn parse_bool_attr(e: &BytesStart<'_>) -> Result<bool> {
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

fn invalid_attribute(description: &str, value: &[u8]) -> Error {
    Error::Invalid(format!(
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

fn required_string_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<String> {
    element
        .try_get_attribute(name)
        .map_err(|error| Error::Xml(error.to_string()))?
        .ok_or_else(|| missing_attribute(description))?
        .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
        .map(|value| value.into_owned())
        .map_err(|error| Error::Xml(error.to_string()))
}

fn optional_u32_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    default: u32,
    description: &str,
) -> Result<u32> {
    let Some(value) = get_attr(element, name) else {
        return Ok(default);
    };
    std::str::from_utf8(&value)
        .ok()
        .and_then(|value| value.parse().ok())
        .ok_or_else(|| invalid_attribute(description, &value))
}

fn optional_i32_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    default: i32,
    description: &str,
) -> Result<i32> {
    let Some(value) = get_attr(element, name) else {
        return Ok(default);
    };
    std::str::from_utf8(&value)
        .ok()
        .and_then(|value| value.parse().ok())
        .ok_or_else(|| invalid_attribute(description, &value))
}

fn optional_bool_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    default: bool,
    description: &str,
) -> Result<bool> {
    match get_attr(element, name) {
        Some(value) => parse_bool_value(&value, description),
        None => Ok(default),
    }
}

fn required_named_f64_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    description: &str,
) -> Result<f64> {
    let value = get_attr(element, name).ok_or_else(|| missing_attribute(description))?;
    std::str::from_utf8(&value)
        .ok()
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite())
        .ok_or_else(|| invalid_attribute(description, &value))
}

fn required_positive_u32_attr(element: &BytesStart<'_>, description: &str) -> Result<u32> {
    let value = required_u32_attr(element, description)?;
    if value == 0 {
        return Err(Error::Invalid(format!("{description} must be positive")));
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
        return Err(Error::Invalid(format!("{description} must be positive")));
    }
    Ok(value)
}

fn required_nonnegative_f64_attr(element: &BytesStart<'_>, description: &str) -> Result<f64> {
    let value = required_f64_attr(element, description)?;
    if value < 0.0 {
        return Err(Error::Invalid(format!("{description} must be nonnegative")));
    }
    Ok(value)
}

fn required_enum_attr(element: &BytesStart<'_>, description: &str) -> Result<String> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute(description))?;
    String::from_utf8(value).map_err(|error| Error::Invalid(error.to_string()))
}

fn bounded_u32_attr(
    element: &BytesStart<'_>,
    description: &str,
    minimum: u32,
    maximum: u32,
) -> Result<u32> {
    let value = required_u32_attr(element, description)?;
    if !(minimum..=maximum).contains(&value) {
        return Err(Error::Invalid(format!(
            "{description} must be between {minimum} and {maximum}"
        )));
    }
    Ok(value)
}

fn bounded_percentage_u32_attr(
    element: &BytesStart<'_>,
    description: &str,
    minimum: u32,
    maximum: u32,
) -> Result<u32> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute(description))?;
    let text = std::str::from_utf8(&value).map_err(|_| invalid_attribute(description, &value))?;
    let digits = text.strip_suffix('%').unwrap_or(text);
    let parsed = digits
        .parse::<u32>()
        .map_err(|_| invalid_attribute(description, &value))?;
    if !(minimum..=maximum).contains(&parsed) {
        return Err(Error::Invalid(format!(
            "{description} must be between {minimum} and {maximum}"
        )));
    }
    Ok(parsed)
}

fn bounded_percentage_i32_attr(
    element: &BytesStart<'_>,
    description: &str,
    minimum: i32,
    maximum: i32,
) -> Result<i32> {
    let value = get_attr(element, b"val").ok_or_else(|| missing_attribute(description))?;
    let text = std::str::from_utf8(&value).map_err(|_| invalid_attribute(description, &value))?;
    let digits = text.strip_suffix('%').unwrap_or(text);
    let parsed = digits
        .parse::<i32>()
        .map_err(|_| invalid_attribute(description, &value))?;
    if !(minimum..=maximum).contains(&parsed) {
        return Err(Error::Invalid(format!(
            "{description} must be between {minimum} and {maximum}"
        )));
    }
    Ok(parsed)
}

fn missing_attribute(description: &str) -> Error {
    Error::Invalid(format!("{description} is missing its value"))
}

#[inline]
fn get_attr(e: &BytesStart<'_>, name: &[u8]) -> Option<Vec<u8>> {
    e.attributes()
        .filter_map(|a| a.ok())
        .find(|a| a.key.as_ref() == name)
        .map(|a| a.value.to_vec())
}
