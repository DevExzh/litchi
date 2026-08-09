//! XML stream façade for chart reading.
//!
//! This layer bounds and preprocesses the input, drives the namespace-aware
//! reader, and dispatches chart elements to the semantic decoder.

use super::super::model::{ChartXmlReader, IGNORED_NAMESPACE_ELEMENT};
use super::semantic::{
    parse_chart_protection, parse_color_map_override, parse_external_data, parse_legend,
    parse_pivot_formats, parse_pivot_source, parse_plot_area, parse_print_settings, parse_title,
    parse_view_3d, parse_wall_floor, required_chart_relationship_id,
};
use super::validation::{
    bounded_u32_attr, parse_bool_attr, parse_display_blanks, required_string_attr,
};

use crate::chart::model::{
    Chart, ExtensionList, ExternalData, PrintSettings, Protection, ShapeProperties, TextProperties,
    UserShapes, View3D, WallFloor,
};
use crate::chart::plot_area::PlotArea;
use crate::{Error, Result};
use quick_xml::events::Event;
use std::io::{BufRead, Read};

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
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
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

pub(super) fn consume_empty_chart_element<R: BufRead>(
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
            Ok(Event::Start(ref element) | Event::Empty(ref element))
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
