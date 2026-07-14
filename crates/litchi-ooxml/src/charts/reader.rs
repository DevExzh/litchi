//! Chart XML reader.
//!
//! This module provides functionality to parse chart XML files
//! from OOXML packages.

use crate::charts::axis::{
    Axis, AxisCommon, AxisCrossBetween, AxisCrossMode, AxisLabelAlign, BuiltInUnit, CategoryAxis,
    DateAxis, DisplayUnits, SeriesAxis, TimeUnit, ValueAxis,
};
use crate::charts::chart::{
    Chart, ChartExtensionList, ChartExternalData, ChartHeaderFooter, ChartPageMargins,
    ChartPageOrientation, ChartPageSetup, ChartPrintSettings, ChartProtection,
    ChartShapeProperties, ChartTextProperties, ChartUserShapes, ColorMapOverride, ColorMapping,
    ColorSchemeIndex, PictureFormat, PictureOptions, PivotFormat, PivotSource, View3D, WallFloor,
};
use crate::charts::legend::{Legend, LegendEntry};
use crate::charts::models::{
    DataSourceRef, Layout, NumberFormat, NumericData, RichText, StringData, TitleText,
};
use crate::charts::plot_area::{
    Area3DTypeGroup, AreaTypeGroup, BandFormat, Bar3DTypeGroup, BarShape, BarTypeGroup,
    BubbleTypeGroup, ChartLines, DataTable, DoughnutTypeGroup, Line3DTypeGroup, LineTypeGroup,
    OfPieTypeGroup, Pie3DTypeGroup, PieTypeGroup, PlotArea, RadarTypeGroup, ScatterTypeGroup,
    StockTypeGroup, Surface3DTypeGroup, SurfaceTypeGroup, TypeGroup, TypeGroupCommon, UpDownBars,
};
use crate::charts::series::{
    DataLabel, DataLabels, DataPoint, ErrorBar, ErrorBarDirection, ErrorBarType, ErrorBarValueType,
    Marker, Series, Trendline, TrendlineType,
};
use crate::charts::types::{
    AxisOrientation, AxisPosition, BarDirection, BarGrouping, DataLabelPosition, DisplayBlanks,
    LayoutMode, LayoutTarget, LegendPosition, MarkerStyle, OfPieSplitType, OfPieType, RadarStyle,
    ScatterStyle, TickLabelPosition, TickMark,
};
use crate::common::xml::{decode_xml_reference, is_drawingml_chart_name, is_drawingml_name};
use crate::error::{OoxmlError, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::{Config, NsReader};
use quick_xml::writer::Writer;
use std::io::BufRead;

const IGNORED_NAMESPACE_ELEMENT: &str = "ignoredNamespaceElement";
const INVALID_COLOR_MAPPING_ELEMENT: &str = "invalidColorMappingElement";

/// Namespace-aware streaming adapter for the chart model parser.
///
/// Core chart elements are exposed unchanged. DrawingML text and color-map choice
/// elements are also kept so their typed models can be decoded, while all other
/// namespaces are skipped as extension content. Rewriting the remaining DrawingML
/// container names prevents them from being mistaken for same-local-name chart
/// elements by the focused parsers below.
struct ChartXmlReader<R: BufRead> {
    inner: NsReader<R>,
    depth: usize,
    skipped_depth: usize,
    saw_root: bool,
    closed_root: bool,
    root_namespace_attributes: Vec<(Vec<u8>, Vec<u8>)>,
}

impl<R: BufRead> ChartXmlReader<R> {
    fn from_reader(reader: R) -> Self {
        Self {
            inner: NsReader::from_reader(reader),
            depth: 0,
            skipped_depth: 0,
            saw_root: false,
            closed_root: false,
            root_namespace_attributes: Vec::new(),
        }
    }

    fn config_mut(&mut self) -> &mut Config {
        self.inner.config_mut()
    }

    fn decoder(&self) -> Decoder {
        self.inner.decoder()
    }

    fn relationship_attribute_value(
        &self,
        element: &BytesStart<'_>,
        name: &[u8],
    ) -> Result<Option<String>> {
        const RELATIONSHIPS_NAMESPACE: &[u8] =
            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
            b"http://purl.oclc.org/ooxml/officeDocument/relationships";

        let mut value = None;
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
            if attribute.key.local_name().as_ref() != name {
                continue;
            }
            let (namespace, _) = self.inner.resolver().resolve_attribute(attribute.key);
            let is_relationship = matches!(
                namespace,
                ResolveResult::Bound(Namespace(value))
                    if value == RELATIONSHIPS_NAMESPACE
                        || value == STRICT_RELATIONSHIPS_NAMESPACE
            ) || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"r");
            if !is_relationship {
                continue;
            }
            if value.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "chart element contains duplicate relationship IDs".into(),
                ));
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Explicit1_0, self.decoder())
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?
                    .into_owned(),
            );
        }
        Ok(value)
    }

    fn make_fragment_root_self_contained(&self, element: &BytesStart<'_>) -> BytesStart<'static> {
        let mut root = element.to_owned();
        let existing_names: Vec<Vec<u8>> = root
            .attributes()
            .filter_map(std::result::Result::ok)
            .map(|attribute| attribute.key.as_ref().to_vec())
            .collect();
        for (name, value) in &self.root_namespace_attributes {
            if !existing_names.iter().any(|existing| existing == name) {
                root.push_attribute((name.as_slice(), value.as_slice()));
            }
        }
        root
    }

    fn capture_empty_fragment(&self, element: &BytesStart<'_>) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        writer
            .write_event(Event::Empty(
                self.make_fragment_root_self_contained(element),
            ))
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        Ok(writer.into_inner())
    }

    fn capture_fragment(&mut self, element: &BytesStart<'_>, description: &str) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        writer
            .write_event(Event::Start(
                self.make_fragment_root_self_contained(element),
            ))
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let fragment_depth = self.depth;
        let mut buffer = Vec::new();
        loop {
            let (_, event) = self
                .inner
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Start(_) => {
                    self.depth = self.depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat(format!("{description} XML nesting is too deep"))
                    })?;
                },
                Event::End(_) => {
                    self.depth = self.depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat(format!(
                            "{description} has an unmatched closing element"
                        ))
                    })?;
                },
                Event::Eof => {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "unterminated {description}"
                    )));
                },
                Event::DocType(_) => {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "{description} cannot contain a document type"
                    )));
                },
                _ => {},
            }
            let finished = matches!(event, Event::End(_)) && self.depth < fragment_depth;
            writer
                .write_event(event.into_owned())
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            buffer.clear();
            if finished {
                break;
            }
        }
        Ok(writer.into_inner())
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
                    self.root_namespace_attributes.clear();
                    for attribute in element.attributes() {
                        let attribute =
                            attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        let name = attribute.key.as_ref();
                        if name == b"xmlns" || name.starts_with(b"xmlns:") {
                            self.root_namespace_attributes
                                .push((name.to_vec(), attribute.value.into_owned()));
                        }
                    }
                    for (name, value) in [
                        (
                            b"xmlns:c".as_slice(),
                            b"http://schemas.openxmlformats.org/drawingml/2006/chart".as_slice(),
                        ),
                        (
                            b"xmlns:a".as_slice(),
                            b"http://schemas.openxmlformats.org/drawingml/2006/main".as_slice(),
                        ),
                        (
                            b"xmlns:r".as_slice(),
                            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                .as_slice(),
                        ),
                    ] {
                        if !self
                            .root_namespace_attributes
                            .iter()
                            .any(|(existing, _)| existing == name)
                        {
                            self.root_namespace_attributes
                                .push((name.to_vec(), value.to_vec()));
                        }
                    }
                }
                self.depth = self.depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("chart XML nesting is too deep".to_string())
                })?;

                if self.skipped_depth > 0 {
                    self.skipped_depth = self.skipped_depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("chart XML nesting is too deep".to_string())
                    })?;
                    element.set_name(IGNORED_NAMESPACE_ELEMENT.as_bytes());
                } else if is_chart && is_drawing_color_map_choice(element.local_name().as_ref()) {
                    element.set_name(INVALID_COLOR_MAPPING_ELEMENT.as_bytes());
                } else if !is_chart && !is_drawing {
                    self.skipped_depth = self.skipped_depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("chart XML nesting is too deep".to_string())
                    })?;
                    element.set_name(IGNORED_NAMESPACE_ELEMENT.as_bytes());
                } else if is_drawing && !is_preserved_drawing_element(element.local_name().as_ref())
                {
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
                if self.skipped_depth > 0 {
                    element.set_name(IGNORED_NAMESPACE_ELEMENT.as_bytes());
                } else if is_chart && is_drawing_color_map_choice(element.local_name().as_ref()) {
                    element.set_name(INVALID_COLOR_MAPPING_ELEMENT.as_bytes());
                } else if !is_chart
                    && (!is_drawing || !is_preserved_drawing_element(element.local_name().as_ref()))
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
                if is_drawing && !is_preserved_drawing_element(element.local_name().as_ref()) {
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

fn is_preserved_drawing_element(local_name: &[u8]) -> bool {
    local_name == b"t" || is_drawing_color_map_choice(local_name)
}

fn is_drawing_color_map_choice(local_name: &[u8]) -> bool {
    matches!(local_name, b"masterClrMapping" | b"overrideClrMapping")
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
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lang" => {
                if chart.language.is_some() {
                    return Err(OoxmlError::InvalidFormat(
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
                    return Err(OoxmlError::InvalidFormat(
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
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate pivot sources".into(),
                    ));
                }
                chart.pivot_source = Some(parse_pivot_source(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"pivotSource" => {
                return Err(OoxmlError::InvalidFormat(
                    "chart pivot source requires a name and format ID".into(),
                ));
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrMapOvr" => {
                if chart.color_map_override.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate color-map overrides".into(),
                    ));
                }
                chart.color_map_override = Some(parse_color_map_override(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"clrMapOvr" => {
                return Err(OoxmlError::InvalidFormat(
                    "chart color-map override requires a mapping choice".into(),
                ));
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"protection" => {
                if chart.protection.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate protection settings".into(),
                    ));
                }
                chart.protection = Some(parse_chart_protection(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"protection" => {
                if chart.protection.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate protection settings".into(),
                    ));
                }
                chart.protection = Some(ChartProtection::default());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pivotFmts" => {
                if chart.pivot_formats.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate pivot-format collections".into(),
                    ));
                }
                chart.pivot_formats = Some(parse_pivot_formats(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"pivotFmts" => {
                if chart.pivot_formats.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate pivot-format collections".into(),
                    ));
                }
                chart.pivot_formats = Some(Vec::new());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"view3D" => {
                if chart.view_3d.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate 3D views".into(),
                    ));
                }
                chart.view_3d = Some(parse_view_3d(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"view3D" => {
                if chart.view_3d.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate 3D views".into(),
                    ));
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
                    _ => unreachable!(),
                };
                if target.is_some() {
                    return Err(OoxmlError::InvalidFormat(
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
                    _ => unreachable!(),
                };
                if target.is_some() {
                    return Err(OoxmlError::InvalidFormat(
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
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate print settings".into(),
                    ));
                }
                chart.print_settings = Some(parse_print_settings(&mut xml_reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"printSettings" => {
                if chart.print_settings.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate print settings".into(),
                    ));
                }
                chart.print_settings = Some(ChartPrintSettings::new());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"externalData" => {
                if chart.external_data.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate external-data relationships".into(),
                    ));
                }
                let relationship_id = required_chart_relationship_id(&xml_reader, e)?;
                chart.external_data = Some(parse_external_data(&mut xml_reader, relationship_id)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"externalData" => {
                if chart.external_data.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate external-data relationships".into(),
                    ));
                }
                chart.external_data = Some(ChartExternalData::new(required_chart_relationship_id(
                    &xml_reader,
                    e,
                )?));
            },
            Ok(Event::Start(ref e)) if closed_chart && e.local_name().as_ref() == b"spPr" => {
                if chart.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate chart-space shape properties".into(),
                    ));
                }
                chart.shape_properties = Some(ChartShapeProperties::from_xml(
                    xml_reader.capture_fragment(e, "chart-space shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if closed_chart && e.local_name().as_ref() == b"spPr" => {
                if chart.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate chart-space shape properties".into(),
                    ));
                }
                chart.shape_properties = Some(ChartShapeProperties::from_xml(
                    xml_reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if closed_chart && e.local_name().as_ref() == b"txPr" => {
                if chart.text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate chart-space text properties".into(),
                    ));
                }
                chart.text_properties = Some(ChartTextProperties::from_xml(
                    xml_reader.capture_fragment(e, "chart-space text properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if closed_chart && e.local_name().as_ref() == b"txPr" => {
                if chart.text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate chart-space text properties".into(),
                    ));
                }
                chart.text_properties = Some(ChartTextProperties::from_xml(
                    xml_reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"userShapes" => {
                if chart.user_shapes.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate user-shapes relationships".into(),
                    ));
                }
                chart.user_shapes = Some(ChartUserShapes::new(required_chart_relationship_id(
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
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate user-shapes relationships".into(),
                    ));
                }
                chart.user_shapes = Some(ChartUserShapes::new(required_chart_relationship_id(
                    &xml_reader,
                    e,
                )?));
            },
            Ok(Event::Start(ref e))
                if saw_chart && !closed_chart && e.local_name().as_ref() == b"extLst" =>
            {
                if chart.chart_extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate chart extension lists".into(),
                    ));
                }
                chart.chart_extension_list = Some(ChartExtensionList::from_xml(
                    xml_reader.capture_fragment(e, "chart extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e))
                if saw_chart && !closed_chart && e.local_name().as_ref() == b"extLst" =>
            {
                if chart.chart_extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate chart extension lists".into(),
                    ));
                }
                chart.chart_extension_list = Some(ChartExtensionList::from_xml(
                    xml_reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if closed_chart && e.local_name().as_ref() == b"extLst" => {
                if chart.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate chart-space extension lists".into(),
                    ));
                }
                chart.extension_list = Some(ChartExtensionList::from_xml(
                    xml_reader.capture_fragment(e, "chart-space extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if closed_chart && e.local_name().as_ref() == b"extLst" => {
                if chart.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart contains duplicate chart-space extension lists".into(),
                    ));
                }
                chart.extension_list = Some(ChartExtensionList::from_xml(
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
                            return Err(OoxmlError::InvalidFormat(
                                "chart contains duplicate titles".into(),
                            ));
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

fn parse_pivot_source<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<PivotSource> {
    let mut name = None;
    let mut format_id = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"name" => {
                if name.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot source contains duplicate names".into(),
                    ));
                }
                name = Some(parse_text_element(reader, b"name")?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"name" => {
                if name.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot source contains duplicate names".into(),
                    ));
                }
                name = Some(String::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"fmtId" => {
                if format_id.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot source contains duplicate format IDs".into(),
                    ));
                }
                format_id = Some(required_u32_attr(element, "chart pivot-source format ID")?);
                consume_empty_chart_element(reader, b"fmtId", "chart pivot-source format ID")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"fmtId" => {
                if format_id.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot source contains duplicate format IDs".into(),
                    ));
                }
                format_id = Some(required_u32_attr(element, "chart pivot-source format ID")?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"pivotSource" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart pivot source".into(),
                ));
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

fn parse_chart_protection<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<ChartProtection> {
    let mut protection = ChartProtection::default();
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
                    return Err(OoxmlError::InvalidFormat(format!(
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
                    return Err(OoxmlError::InvalidFormat(format!(
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
                return Err(OoxmlError::InvalidFormat(
                    "chart protection is not closed".into(),
                ));
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
                    return Err(OoxmlError::InvalidFormat(
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
                    return Err(OoxmlError::InvalidFormat(
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
                return Err(OoxmlError::InvalidFormat(
                    "chart color-map override contains an unexpected choice".into(),
                ));
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"clrMapOvr" => {
                return mapping.ok_or_else(|| {
                    OoxmlError::InvalidFormat(
                        "chart color-map override requires a mapping choice".into(),
                    )
                });
            },
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
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
        return Err(OoxmlError::InvalidFormat(
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
        .ok_or_else(|| OoxmlError::InvalidFormat("chart relationship ID is required".into()))
}

fn parse_external_data<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    relationship_id: String,
) -> Result<ChartExternalData> {
    let mut external_data = ChartExternalData::new(relationship_id);
    let mut saw_auto_update = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"autoUpdate" => {
                if saw_auto_update {
                    return Err(OoxmlError::InvalidFormat(
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
                    return Err(OoxmlError::InvalidFormat(
                        "chart external data contains duplicate auto-update settings".into(),
                    ));
                }
                saw_auto_update = true;
                external_data.auto_update = Some(parse_bool_attr(element)?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() != IGNORED_NAMESPACE_ELEMENT.as_bytes() =>
            {
                return Err(OoxmlError::InvalidFormat(
                    "chart external data contains an unexpected child".into(),
                ));
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"externalData" => {
                return Ok(external_data);
            },
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "chart external data is not closed".into(),
                ));
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
        OoxmlError::InvalidFormat(format!(
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
                    return Err(OoxmlError::InvalidFormat(format!(
                        "chart contains duplicate pivot-format index {}",
                        format.index
                    )));
                }
                formats.push(format);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"pivotFmt" => {
                return Err(OoxmlError::InvalidFormat(
                    "chart pivot format is missing its index".into(),
                ));
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"pivotFmts" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart pivot formats".into(),
                ));
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
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot format contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart pivot-format shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot format contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot format contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_fragment(element, "chart pivot-format text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot format contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"marker" => {
                if marker.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot format contains duplicate markers".into(),
                    ));
                }
                marker = Some(parse_series_marker(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"marker" => {
                if marker.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot format contains duplicate markers".into(),
                    ));
                }
                marker = Some(Marker::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"dLbl" => {
                if data_label.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot format contains duplicate data labels".into(),
                    ));
                }
                data_label = Some(parse_data_label(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dLbl" => {
                return Err(OoxmlError::InvalidFormat(
                    "chart pivot-format data label is missing its index".into(),
                ));
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"idx" =>
            {
                if index.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot format contains duplicate indexes".into(),
                    ));
                }
                index = Some(required_u32_attr(element, "chart pivot-format index")?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot format contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(element, "chart pivot-format extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart pivot format contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"pivotFmt" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart pivot format".into(),
                ));
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

fn parse_print_settings<R: BufRead>(reader: &mut ChartXmlReader<R>) -> Result<ChartPrintSettings> {
    let mut settings = ChartPrintSettings::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"headerFooter" => {
                if settings.header_footer.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart print settings contain duplicate headers and footers".into(),
                    ));
                }
                settings.header_footer = Some(parse_chart_header_footer(reader, element)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"headerFooter" => {
                if settings.header_footer.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart print settings contain duplicate headers and footers".into(),
                    ));
                }
                settings.header_footer = Some(parse_chart_header_footer_attributes(element)?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"pageMargins" => {
                if settings.page_margins.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart print settings contain duplicate page margins".into(),
                    ));
                }
                settings.page_margins = Some(parse_chart_page_margins(element)?);
                consume_empty_chart_element(reader, b"pageMargins", "chart page margins")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"pageMargins" => {
                if settings.page_margins.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart print settings contain duplicate page margins".into(),
                    ));
                }
                settings.page_margins = Some(parse_chart_page_margins(element)?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"pageSetup" => {
                if settings.page_setup.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart print settings contain duplicate page setup".into(),
                    ));
                }
                settings.page_setup = Some(parse_chart_page_setup(element)?);
                consume_empty_chart_element(reader, b"pageSetup", "chart page setup")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"pageSetup" => {
                if settings.page_setup.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart print settings contain duplicate page setup".into(),
                    ));
                }
                settings.page_setup = Some(parse_chart_page_setup(element)?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"printSettings" => {
                break;
            },
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart print settings".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(settings)
}

fn parse_chart_header_footer_attributes(element: &BytesStart<'_>) -> Result<ChartHeaderFooter> {
    let mut header_footer = ChartHeaderFooter::new();
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
) -> Result<ChartHeaderFooter> {
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
                        return Err(OoxmlError::InvalidFormat(
                            "chart header/footer contains an unexpected child".into(),
                        ));
                    },
                };
                if value.is_some() {
                    return Err(OoxmlError::InvalidFormat(
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
                        return Err(OoxmlError::InvalidFormat(
                            "chart header/footer contains an unexpected child".into(),
                        ));
                    },
                };
                if value.replace(String::new()).is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart header/footer contains a duplicate string".into(),
                    ));
                }
            },
            Ok(Event::End(ref child)) if child.local_name().as_ref() == b"headerFooter" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart header/footer".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(header_footer)
}

fn parse_chart_page_margins(element: &BytesStart<'_>) -> Result<ChartPageMargins> {
    Ok(ChartPageMargins::new(
        required_named_f64_attr(element, b"l", "chart left page margin")?,
        required_named_f64_attr(element, b"r", "chart right page margin")?,
        required_named_f64_attr(element, b"t", "chart top page margin")?,
        required_named_f64_attr(element, b"b", "chart bottom page margin")?,
        required_named_f64_attr(element, b"header", "chart header page margin")?,
        required_named_f64_attr(element, b"footer", "chart footer page margin")?,
    ))
}

fn parse_chart_page_setup(element: &BytesStart<'_>) -> Result<ChartPageSetup> {
    let mut setup = ChartPageSetup::new();
    setup.paper_size = optional_u32_attr(element, b"paperSize", 1, "chart printer paper size")?;
    setup.first_page_number =
        optional_u32_attr(element, b"firstPageNumber", 1, "chart first page number")?;
    setup.orientation = match get_attr(element, b"orientation").as_deref() {
        None | Some(b"default") => ChartPageOrientation::Default,
        Some(b"portrait") => ChartPageOrientation::Portrait,
        Some(b"landscape") => ChartPageOrientation::Landscape,
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
                return Err(OoxmlError::InvalidFormat(format!(
                    "{description} contains child elements"
                )));
            },
            Ok(Event::Text(ref text))
                if !text
                    .decode()
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?
                    .trim()
                    .is_empty() =>
            {
                return Err(OoxmlError::InvalidFormat(format!(
                    "{description} contains text"
                )));
            },
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(format!(
                    "unterminated {description}"
                )));
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
    shape_properties: Option<ChartShapeProperties>,
    text_properties: Option<ChartTextProperties>,
    extension_list: Option<ChartExtensionList>,
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
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart title contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart title shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart title contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart title contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_fragment(element, "chart title text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart title contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart title contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(element, "chart title extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart title contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
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
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"thickness" => {
                if wall_floor.thickness.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart surface contains duplicate thickness values".into(),
                    ));
                }
                wall_floor.thickness = Some(bounded_u32_attr(e, "chart wall thickness", 0, 4096)?);
                consume_empty_chart_element(reader, b"thickness", "chart surface thickness")?;
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"thickness" => {
                if wall_floor.thickness.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart surface contains duplicate thickness values".into(),
                    ));
                }
                wall_floor.thickness = Some(bounded_u32_attr(e, "chart wall thickness", 0, 4096)?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if wall_floor.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart surface contains duplicate shape properties".into(),
                    ));
                }
                wall_floor.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart surface shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if wall_floor.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart surface contains duplicate shape properties".into(),
                    ));
                }
                wall_floor.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pictureOptions" => {
                if wall_floor.picture_options.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart surface contains duplicate picture options".into(),
                    ));
                }
                wall_floor.picture_options = Some(parse_picture_options(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"pictureOptions" => {
                if wall_floor.picture_options.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart surface contains duplicate picture options".into(),
                    ));
                }
                wall_floor.picture_options = Some(PictureOptions::default());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if wall_floor.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart surface contains duplicate extension lists".into(),
                    ));
                }
                wall_floor.extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(e, "chart surface extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if wall_floor.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart surface contains duplicate extension lists".into(),
                    ));
                }
                wall_floor.extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
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
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart picture options".into(),
                ));
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
                return Err(OoxmlError::InvalidFormat(
                    "chart picture options contain duplicate formats".into(),
                ));
            }
            let value = required_string_attr(element, b"val", decoder, "chart picture format")?;
            options.picture_format = Some(match value.as_str() {
                "stretch" => PictureFormat::Stretch,
                "stack" => PictureFormat::Stack,
                "stackScale" => PictureFormat::StackScale,
                _ => {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "invalid chart picture format '{value}'"
                    )));
                },
            });
            None
        },
        b"pictureStackUnit" => {
            if options.picture_stack_unit.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "chart picture options contain duplicate stack units".into(),
                ));
            }
            options.picture_stack_unit =
                Some(required_f64_attr(element, "chart picture stack unit")?);
            None
        },
        _ => unreachable!("picture option child was checked by caller"),
    };
    if let Some(target) = target {
        if target.is_some() {
            return Err(OoxmlError::InvalidFormat(
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
                    return Err(OoxmlError::InvalidFormat(
                        "chart plot area contains duplicate data tables".into(),
                    ));
                }
                saw_data_table = true;
                plot_area.data_table = Some(parse_data_table(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dTable" => {
                if saw_data_table {
                    return Err(OoxmlError::InvalidFormat(
                        "chart plot area contains duplicate data tables".into(),
                    ));
                }
                saw_data_table = true;
                plot_area.data_table = Some(DataTable::default());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if plot_area.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart plot area contains duplicate shape properties".into(),
                    ));
                }
                plot_area.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart plot-area shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if plot_area.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart plot area contains duplicate shape properties".into(),
                    ));
                }
                plot_area.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if plot_area.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart plot area contains duplicate extension lists".into(),
                    ));
                }
                plot_area.extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(e, "chart plot-area extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if plot_area.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart plot area contains duplicate extension lists".into(),
                    ));
                }
                plot_area.extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Empty(ref e))
                if is_chart_type_group_name(e.local_name().as_ref())
                    || matches!(
                        e.local_name().as_ref(),
                        b"catAx" | b"valAx" | b"dateAx" | b"serAx"
                    ) =>
            {
                return Err(OoxmlError::InvalidFormat(format!(
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
                    return Err(OoxmlError::InvalidFormat(
                        "chart data table contains duplicate shape properties".into(),
                    ));
                }
                data_table.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart data-table shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if data_table.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data table contains duplicate shape properties".into(),
                    ));
                }
                data_table.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if data_table.text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data table contains duplicate text properties".into(),
                    ));
                }
                data_table.text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_fragment(element, "chart data-table text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if data_table.text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data table contains duplicate text properties".into(),
                    ));
                }
                data_table.text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if data_table.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data table contains duplicate extension lists".into(),
                    ));
                }
                data_table.extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(element, "chart data-table extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if data_table.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data table contains duplicate extension lists".into(),
                    ));
                }
                data_table.extension_list = Some(ChartExtensionList::from_xml(
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
                        return Err(OoxmlError::InvalidFormat(
                            "chart data table contains a duplicate visibility setting".into(),
                        ));
                    }
                    seen[index] = true;
                    *field = parse_bool_attr(element)?;
                }
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"dTable" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart data table".into(),
                ));
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
    supports_data_labels: bool,
    supports_axes: bool,
    mut drop_lines: Option<&mut Option<ChartLines>>,
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
                    set_chart_lines(target, ChartLines::new(), "chart drop lines")?;
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

fn parse_type_group_extension<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    common: &mut TypeGroupCommon,
    element: &BytesStart<'_>,
    empty: bool,
) -> Result<()> {
    if common.extension_list.is_some() {
        return Err(OoxmlError::InvalidFormat(
            "chart type group contains duplicate extension lists".into(),
        ));
    }
    let xml = if empty {
        reader.capture_empty_fragment(element)?
    } else {
        reader.capture_fragment(element, "chart type-group extension list")?
    };
    common.extension_list = Some(ChartExtensionList::from_xml(xml)?);
    Ok(())
}

fn begin_group_data_labels(seen: &mut bool) -> Result<()> {
    if *seen {
        return Err(OoxmlError::InvalidFormat(
            "chart type group contains duplicate data-label settings".into(),
        ));
    }
    *seen = true;
    Ok(())
}

fn set_chart_lines(
    target: &mut Option<ChartLines>,
    lines: ChartLines,
    description: &str,
) -> Result<()> {
    if target.replace(lines).is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "{description} are duplicated"
        )));
    }
    Ok(())
}

fn parse_chart_lines<R: BufRead>(
    reader: &mut ChartXmlReader<R>,
    end_name: &[u8],
) -> Result<ChartLines> {
    let mut lines = ChartLines::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if lines.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart lines contain duplicate shape properties".into(),
                    ));
                }
                lines.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart-line shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if lines.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart lines contain duplicate shape properties".into(),
                    ));
                }
                lines.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == end_name => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart-line formatting".into(),
                ));
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
        return Err(OoxmlError::InvalidFormat(format!(
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
    let mut bubble_3d = false;
    let mut bubble_scale = None;
    let mut show_negative_bubbles = true;
    let mut size_represents = "area".to_string();
    let common = parse_common_type_group(reader, b"bubbleChart", true, true, None, |element| {
        match element.local_name().as_ref() {
            b"bubble3D" => bubble_3d = parse_bool_attr(element)?,
            b"bubbleScale" => {
                bubble_scale = Some(match get_attr(element, b"val") {
                    Some(_) => bounded_percentage_u32_attr(element, "bubble scale", 0, 300)?,
                    None => 100,
                });
            },
            b"showNegBubbles" => show_negative_bubbles = parse_bool_attr(element)?,
            b"sizeRepresents" => {
                let value = get_attr(element, b"val").unwrap_or_else(|| b"area".to_vec());
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
    let common = parse_common_type_group(reader, b"doughnutChart", true, false, None, |element| {
        match element.local_name().as_ref() {
            b"firstSliceAng" => {
                first_slice_angle = match get_attr(element, b"val") {
                    Some(_) => required_u32_attr(element, "first-slice angle")?,
                    None => 0,
                };
                if first_slice_angle > 360 {
                    return Err(OoxmlError::InvalidFormat(
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
                group.series_lines.push(ChartLines::new());
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
                    return Err(OoxmlError::InvalidFormat(
                        "of-pie chart contains duplicate custom splits".into(),
                    ));
                }
                saw_custom_split = true;
                group.custom_split_points = Some(parse_custom_pie_split(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"custSplit" => {
                if saw_custom_split {
                    return Err(OoxmlError::InvalidFormat(
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
                            return Err(OoxmlError::InvalidFormat(
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
                            return Err(OoxmlError::InvalidFormat(
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
                            return Err(OoxmlError::InvalidFormat(
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
                            return Err(OoxmlError::InvalidFormat(
                                "of-pie chart contains duplicate split positions".into(),
                            ));
                        }
                        saw_split_position = true;
                        group.split_position =
                            Some(required_f64_attr(element, "of-pie split position")?);
                    },
                    b"secondPieSize" => {
                        if saw_second_pie_size {
                            return Err(OoxmlError::InvalidFormat(
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
                return Err(OoxmlError::InvalidFormat(
                    "unterminated of-pie chart".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }

    if !saw_of_pie_type {
        return Err(OoxmlError::InvalidFormat(
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
                return Err(OoxmlError::InvalidFormat(
                    "unterminated of-pie custom split".into(),
                ));
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
                set_chart_lines(&mut bars.up_bars, ChartLines::new(), "chart up bars")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"downBars" => {
                let lines = parse_chart_lines(reader, b"downBars")?;
                set_chart_lines(&mut bars.down_bars, lines, "chart down bars")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"downBars" => {
                set_chart_lines(&mut bars.down_bars, ChartLines::new(), "chart down bars")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if bars.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart up/down bars contain duplicate extension lists".into(),
                    ));
                }
                bars.extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(element, "chart up/down-bar extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if bars.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart up/down bars contain duplicate extension lists".into(),
                    ));
                }
                bars.extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"gapWidth" =>
            {
                if saw_gap_width {
                    return Err(OoxmlError::InvalidFormat(
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
                return Err(OoxmlError::InvalidFormat(
                    "unterminated chart up/down bars".into(),
                ));
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
                set_chart_lines(&mut drop_lines, ChartLines::new(), "stock chart drop lines")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"hiLowLines" => {
                let lines = parse_chart_lines(reader, b"hiLowLines")?;
                set_chart_lines(&mut high_low_lines, lines, "stock chart high/low lines")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"hiLowLines" => {
                set_chart_lines(
                    &mut high_low_lines,
                    ChartLines::new(),
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
                    return Err(OoxmlError::InvalidFormat(
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
                return Err(OoxmlError::InvalidFormat("unterminated stock chart".into()));
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
                    return Err(OoxmlError::InvalidFormat(
                        "surface chart contains duplicate band-format collections".into(),
                    ));
                }
                band_formats = Some(parse_surface_band_formats(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"bandFmts" => {
                if band_formats.is_some() {
                    return Err(OoxmlError::InvalidFormat(
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
                return Err(OoxmlError::InvalidFormat(
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
                return Err(OoxmlError::InvalidFormat(
                    "unterminated surface chart".into(),
                ));
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
                    return Err(OoxmlError::InvalidFormat(format!(
                        "surface chart contains duplicate band index {}",
                        format.index
                    )));
                }
                formats.push(format);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"bandFmt" => {
                return Err(OoxmlError::InvalidFormat(
                    "surface chart band format is missing its index".into(),
                ));
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"bandFmts" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
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
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element))
                if element.local_name().as_ref() == b"idx" =>
            {
                if index.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "surface chart band format contains duplicate indexes".into(),
                    ));
                }
                index = Some(required_u32_attr(element, "surface chart band index")?);
            },
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"bandFmt" => break,
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated surface chart band format".into(),
                ));
            },
            Err(error) => return Err(error),
            _ => {},
        }
        buf.clear();
    }
    Ok(BandFormat::new(index.ok_or_else(|| {
        missing_attribute("surface chart band index")
    })?))
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
                series_lines.push(ChartLines::new());
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
                set_chart_lines(&mut drop_lines, ChartLines::new(), "line chart drop lines")?;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"hiLowLines" => {
                let lines = parse_chart_lines(reader, b"hiLowLines")?;
                set_chart_lines(&mut high_low_lines, lines, "line chart high/low lines")?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"hiLowLines" => {
                set_chart_lines(
                    &mut high_low_lines,
                    ChartLines::new(),
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
                    return Err(OoxmlError::InvalidFormat(
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
                set_chart_lines(&mut drop_lines, ChartLines::new(), "area chart drop lines")?;
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
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if series.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart series contains duplicate shape properties".into(),
                    ));
                }
                series.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_fragment(e, "chart series shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"spPr" => {
                if series.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart series contains duplicate shape properties".into(),
                    ));
                }
                series.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pictureOptions" => {
                if series.picture_options.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart series contains duplicate picture options".into(),
                    ));
                }
                series.picture_options = Some(parse_picture_options(reader)?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"pictureOptions" => {
                if series.picture_options.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart series contains duplicate picture options".into(),
                    ));
                }
                series.picture_options = Some(PictureOptions::default());
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if series.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart series contains duplicate extension lists".into(),
                    ));
                }
                series.extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(e, "chart series extension list")?,
                )?);
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"extLst" => {
                if series.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart series contains duplicate extension lists".into(),
                    ));
                }
                series.extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_empty_fragment(e)?,
                )?);
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"marker" => {
                if saw_marker {
                    return Err(OoxmlError::InvalidFormat(
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
                    return Err(OoxmlError::InvalidFormat(
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
                    b"shape" => {
                        if series.bar_shape.is_some() {
                            return Err(OoxmlError::InvalidFormat(
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
                    return Err(OoxmlError::InvalidFormat(
                        "chart marker has duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart marker shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart marker has duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart marker has duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(element, "chart marker extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart marker has duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
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
                    return Err(OoxmlError::InvalidFormat(
                        "chart data point contains duplicate markers".into(),
                    ));
                }
                marker = Some(parse_series_marker(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"marker" => {
                if marker.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data point contains duplicate markers".into(),
                    ));
                }
                marker = Some(Marker::new());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data point contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart data-point shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data point contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"pictureOptions" => {
                if picture_options.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data point contains duplicate picture options".into(),
                    ));
                }
                picture_options = Some(parse_picture_options(reader)?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"pictureOptions" => {
                if picture_options.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data point contains duplicate picture options".into(),
                    ));
                }
                picture_options = Some(PictureOptions::default());
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data point contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(element, "chart data-point extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data point contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
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
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if labels.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data labels contain duplicate shape properties".into(),
                    ));
                }
                labels.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart data-label shape properties")?,
                )?);
                saw_shared_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if labels.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data labels contain duplicate shape properties".into(),
                    ));
                }
                labels.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
                saw_shared_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if labels.text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data labels contain duplicate text properties".into(),
                    ));
                }
                labels.text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_fragment(element, "chart data-label text properties")?,
                )?);
                saw_shared_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if labels.text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data labels contain duplicate text properties".into(),
                    ));
                }
                labels.text_properties = Some(ChartTextProperties::from_xml(
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
                    ChartLines::new(),
                    "chart data-label leader lines",
                )?;
                saw_shared_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if labels.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data labels contain duplicate extension lists".into(),
                    ));
                }
                labels.extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(element, "chart data-label extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if labels.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart data labels contain duplicate extension lists".into(),
                    ));
                }
                labels.extension_list = Some(ChartExtensionList::from_xml(
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
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if label.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart point data label contains duplicate shape properties".into(),
                    ));
                }
                label.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart point data-label shape properties")?,
                )?);
                saw_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if label.shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart point data label contains duplicate shape properties".into(),
                    ));
                }
                label.shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if label.text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart point data label contains duplicate text properties".into(),
                    ));
                }
                label.text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_fragment(element, "chart point data-label text properties")?,
                )?);
                saw_settings = true;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if label.text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart point data label contains duplicate text properties".into(),
                    ));
                }
                label.text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
                saw_settings = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if label.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart point data label contains duplicate extension lists".into(),
                    ));
                }
                label.extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(element, "chart point data-label extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if label.extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart point data label contains duplicate extension lists".into(),
                    ));
                }
                label.extension_list = Some(ChartExtensionList::from_xml(
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
            return Err(OoxmlError::InvalidFormat(format!(
                "{description} is specified more than once"
            )));
        }
        let value = required_f64_attr(element, description)?;
        if element.local_name().as_ref() == b"logBase" && !(2.0..=1000.0).contains(&value) {
            return Err(OoxmlError::InvalidFormat(
                "chart logarithmic base must be between 2 and 1000".into(),
            ));
        }
        *slot = Some(value);
        Ok(true)
    }

    fn validate(&self) -> Result<()> {
        if self.min.zip(self.max).is_some_and(|(min, max)| min > max) {
            return Err(OoxmlError::InvalidFormat(
                "chart axis minimum exceeds maximum".into(),
            ));
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
    title_shape_properties: Option<ChartShapeProperties>,
    title_text_properties: Option<ChartTextProperties>,
    title_extension_list: Option<ChartExtensionList>,
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
    major_gridlines: Option<ChartLines>,
    minor_gridlines: Option<ChartLines>,
    shape_properties: Option<ChartShapeProperties>,
    text_properties: Option<ChartTextProperties>,
    scaling_extension_list: Option<ChartExtensionList>,
    extension_list: Option<ChartExtensionList>,
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
        return Err(OoxmlError::InvalidFormat(
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
                ChartLines::new()
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
                ChartLines::new()
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
                return Err(OoxmlError::InvalidFormat(
                    "chart axis contains duplicate shape properties".into(),
                ));
            }
            let xml = if empty {
                reader.capture_empty_fragment(element)?
            } else {
                reader.capture_fragment(element, "chart axis shape properties")?
            };
            common.shape_properties = Some(ChartShapeProperties::from_xml(xml)?);
        },
        b"txPr" => {
            if common.text_properties.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "chart axis contains duplicate text properties".into(),
                ));
            }
            let xml = if empty {
                reader.capture_empty_fragment(element)?
            } else {
                reader.capture_fragment(element, "chart axis text properties")?
            };
            common.text_properties = Some(ChartTextProperties::from_xml(xml)?);
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
        return Err(OoxmlError::InvalidFormat(format!(
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
    *target = Some(ChartExtensionList::from_xml(xml)?);
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
                    return Err(OoxmlError::InvalidFormat(
                        "chart display units contain duplicate labels".into(),
                    ));
                }
                saw_label = true;
                parsed_label = parse_display_units_label(reader)?;
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"dispUnitsLbl" => {
                if saw_label {
                    return Err(OoxmlError::InvalidFormat(
                        "chart display units contain duplicate labels".into(),
                    ));
                }
                saw_label = true;
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart display units contain duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(element, "chart display-units extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart display units contain duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
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
    shape_properties: Option<ChartShapeProperties>,
    text_properties: Option<ChartTextProperties>,
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
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart display-units label contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ChartShapeProperties::from_xml(
                    reader
                        .capture_fragment(element, "chart display-units label shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart display-units label contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart display-units label contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(ChartTextProperties::from_xml(
                    reader
                        .capture_fragment(element, "chart display-units label text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart display-units label contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
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

fn unterminated_axis(kind: &str) -> OoxmlError {
    OoxmlError::InvalidFormat(format!("unterminated chart {kind} axis"))
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
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart legend contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_fragment(element, "chart legend shape properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"spPr" => {
                if shape_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart legend contains duplicate shape properties".into(),
                    ));
                }
                shape_properties = Some(ChartShapeProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart legend contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_fragment(element, "chart legend text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart legend contains duplicate text properties".into(),
                    ));
                }
                text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart legend contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(element, "chart legend extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart legend contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
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
                    return Err(OoxmlError::InvalidFormat(
                        "chart legend entry contains multiple choice values".into(),
                    ));
                }
                text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_fragment(element, "chart legend-entry text properties")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"txPr" => {
                if saw_delete || text_properties.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart legend entry contains multiple choice values".into(),
                    ));
                }
                text_properties = Some(ChartTextProperties::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
            Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart legend entry contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_fragment(element, "chart legend-entry extension list")?,
                )?);
            },
            Ok(Event::Empty(ref element)) if element.local_name().as_ref() == b"extLst" => {
                if extension_list.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "chart legend entry contains duplicate extension lists".into(),
                    ));
                }
                extension_list = Some(ChartExtensionList::from_xml(
                    reader.capture_empty_fragment(element)?,
                )?);
            },
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
                        if saw_delete || text_properties.is_some() {
                            return Err(OoxmlError::InvalidFormat(
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
    if !saw_delete && text_properties.is_none() {
        return Err(OoxmlError::InvalidFormat(
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

fn required_string_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<String> {
    element
        .try_get_attribute(name)
        .map_err(|error| OoxmlError::Xml(error.to_string()))?
        .ok_or_else(|| missing_attribute(description))?
        .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
        .map(|value| value.into_owned())
        .map_err(|error| OoxmlError::Xml(error.to_string()))
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
        return Err(OoxmlError::InvalidFormat(format!(
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
        return Err(OoxmlError::InvalidFormat(format!(
            "{description} must be between {minimum} and {maximum}"
        )));
    }
    Ok(parsed)
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
    fn round_trips_and_validates_chart_language_and_pivot_source() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:lang val="zh-Hant"/><c:pivotSource><c:name>Pivot &amp; One</c:name>
                <c:fmtId val="42"/></c:pivotSource>
            <c:chart><c:plotArea/></c:chart></c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        assert_eq!(chart.language.as_deref(), Some("zh-Hant"));
        let source = chart.pivot_source.as_ref().unwrap();
        assert_eq!(source.name, "Pivot & One");
        assert_eq!(source.format_id, 42);

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        assert_eq!(reparsed.language.as_deref(), Some("zh-Hant"));
        assert_eq!(reparsed.pivot_source.as_ref().unwrap().name, "Pivot & One");

        for invalid in [
            br#"<c:lang val="en-US"/><c:lang val="fr-FR"/>"#.as_slice(),
            br#"<c:pivotSource><c:fmtId val="1"/></c:pivotSource>"#.as_slice(),
            br#"<c:pivotSource><c:name>Pivot</c:name></c:pivotSource>"#.as_slice(),
            br#"<c:pivotSource/>"#.as_slice(),
        ] {
            let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">"#.to_vec();
            document.extend_from_slice(invalid);
            document.extend_from_slice(b"<c:chart><c:plotArea/></c:chart></c:chartSpace>");
            assert!(parse_chart(document.as_slice()).is_err());
        }
    }

    #[test]
    fn round_trips_and_validates_chart_protection() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:protection><c:chartObject/><c:data val="0"/><c:formatting val="true"/>
                <c:selection val="false"/><c:userInterface/></c:protection>
            <c:chart><c:plotArea/></c:chart></c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let protection = chart.protection.as_ref().unwrap();
        assert_eq!(protection.chart_object, Some(true));
        assert_eq!(protection.data, Some(false));
        assert_eq!(protection.formatting, Some(true));
        assert_eq!(protection.selection, Some(false));
        assert_eq!(protection.user_interface, Some(true));

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        let protection = reparsed.protection.as_ref().unwrap();
        assert_eq!(protection.chart_object, Some(true));
        assert_eq!(protection.selection, Some(false));

        let empty =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:protection/><c:chart><c:plotArea/></c:chart></c:chartSpace>"#;
        let empty = parse_chart(empty.as_slice()).unwrap();
        let empty = empty.protection.as_ref().unwrap();
        assert_eq!(empty.chart_object, None);
        assert_eq!(empty.user_interface, None);

        for invalid in [
            br#"<c:protection><c:data/><c:data val="0"/></c:protection>"#.as_slice(),
            br#"<c:protection><c:selection val="maybe"/></c:protection>"#.as_slice(),
            br#"<c:protection/><c:protection/>"#.as_slice(),
            br#"<c:protection><c:data><c:data/></c:data></c:protection>"#.as_slice(),
        ] {
            let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">"#.to_vec();
            document.extend_from_slice(invalid);
            document.extend_from_slice(b"<c:chart><c:plotArea/></c:chart></c:chartSpace>");
            assert!(parse_chart(document.as_slice()).is_err());
        }
    }

    #[test]
    fn round_trips_and_validates_chart_color_map_overrides() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="accent1"
                tx2="accent2" accent1="accent3" accent2="accent4" accent3="accent5"
                accent4="accent6" accent5="hlink" accent6="folHlink" hlink="dk2"
                folHlink="lt2"/></c:clrMapOvr>
            <c:chart><c:plotArea/></c:chart></c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let ColorMapOverride::Override(mapping) = chart.color_map_override.as_ref().unwrap() else {
            panic!("expected explicit chart color mapping");
        };
        assert_eq!(mapping.background1, ColorSchemeIndex::Dark1);
        assert_eq!(mapping.background2, ColorSchemeIndex::Accent1);
        assert_eq!(mapping.accent5, ColorSchemeIndex::Hyperlink);
        assert_eq!(mapping.followed_hyperlink, ColorSchemeIndex::Light2);

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        assert_eq!(reparsed.color_map_override, chart.color_map_override);

        let master = br#"<c:chartSpace xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart"
                xmlns:d="http://purl.oclc.org/ooxml/drawingml/main">
            <c:clrMapOvr><d:masterClrMapping></d:masterClrMapping></c:clrMapOvr>
            <c:chart><c:plotArea/></c:chart></c:chartSpace>"#;
        assert_eq!(
            parse_chart(master.as_slice()).unwrap().color_map_override,
            Some(ColorMapOverride::Master)
        );

        for invalid in [
            br#"<c:clrMapOvr/>"#.as_slice(),
            br#"<c:clrMapOvr><a:masterClrMapping/><a:masterClrMapping/></c:clrMapOvr>"#.as_slice(),
            br#"<c:clrMapOvr><a:overrideClrMapping bg1="lt1"/></c:clrMapOvr>"#.as_slice(),
            br#"<c:clrMapOvr><a:overrideClrMapping bg1="none" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></c:clrMapOvr>"#.as_slice(),
            br#"<c:clrMapOvr><c:masterClrMapping/></c:clrMapOvr>"#.as_slice(),
        ] {
            let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#.to_vec();
            document.extend_from_slice(invalid);
            document.extend_from_slice(b"<c:chart><c:plotArea/></c:chart></c:chartSpace>");
            assert!(parse_chart(document.as_slice()).is_err());
        }
    }

    #[test]
    fn round_trips_and_validates_chart_external_data() {
        let xml = br#"<c:chartSpace
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <c:chart><c:plotArea/></c:chart>
            <c:externalData rel:id="rId7"><c:autoUpdate val="0"/></c:externalData>
        </c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let external_data = chart.external_data.as_ref().unwrap();
        assert_eq!(external_data.relationship_id.as_deref(), Some("rId7"));
        assert_eq!(external_data.auto_update, Some(false));

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        assert_eq!(reparsed.external_data, chart.external_data);

        let implicit_true = br#"<c:chartSpace
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <c:chart></c:chart><c:externalData r:id="rId1"><c:autoUpdate/></c:externalData>
        </c:chartSpace>"#;
        assert_eq!(
            parse_chart(implicit_true.as_slice())
                .unwrap()
                .external_data
                .unwrap()
                .auto_update,
            Some(true)
        );

        for invalid in [
            br#"<c:externalData/>"#.as_slice(),
            br#"<c:externalData id="rId1"/>"#.as_slice(),
            br#"<c:externalData r:id="rId1"><c:autoUpdate/><c:autoUpdate/></c:externalData>"#
                .as_slice(),
            br#"<c:externalData r:id="rId1"><c:autoUpdate val="maybe"/></c:externalData>"#
                .as_slice(),
            br#"<c:externalData r:id="rId1"/><c:externalData r:id="rId2"/>"#.as_slice(),
        ] {
            let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart></c:chart>"#.to_vec();
            document.extend_from_slice(invalid);
            document.extend_from_slice(b"</c:chartSpace>");
            assert!(parse_chart(document.as_slice()).is_err());
        }

        let mut pending = Chart::new();
        pending.external_data = Some(ChartExternalData::pending());
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &pending).is_err());
    }

    #[test]
    fn round_trips_and_validates_chart_user_shapes_relationships() {
        let xml = br#"<c:chartSpace
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <c:chart><c:plotArea/></c:chart><c:userShapes q:id="shapeRel"/>
        </c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        assert_eq!(
            chart
                .user_shapes
                .as_ref()
                .unwrap()
                .relationship_id
                .as_deref(),
            Some("shapeRel")
        );
        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        assert_eq!(
            parse_chart(output.as_slice()).unwrap().user_shapes,
            chart.user_shapes
        );

        for invalid in [
            br#"<c:userShapes/>"#.as_slice(),
            br#"<c:userShapes r:id="one"><c:autoUpdate/></c:userShapes>"#.as_slice(),
            br#"<c:userShapes r:id="one"/><c:userShapes r:id="two"/>"#.as_slice(),
        ] {
            let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart></c:chart>"#.to_vec();
            document.extend_from_slice(invalid);
            document.extend_from_slice(b"</c:chartSpace>");
            assert!(parse_chart(document.as_slice()).is_err());
        }
    }

    #[test]
    fn preserves_chart_space_drawing_and_extension_fragments() {
        let xml = br#"<c:chartSpace
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:x="urn:example:chart-extension">
            <c:chart><c:plotArea>
                <c:spPr><a:solidFill><a:srgbClr val="654321"/></a:solidFill></c:spPr>
                <c:extLst><c:ext uri="plot"><x:plotPayload/></c:ext></c:extLst>
            </c:plotArea><c:extLst><c:ext uri="chart"><x:chartPayload/></c:ext></c:extLst></c:chart>
            <c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></c:spPr>
            <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Label</a:t></a:r></a:p></c:txPr>
            <c:extLst><c:ext uri="example"><x:payload enabled="1"/></c:ext></c:extLst>
        </c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let shape_properties = chart.shape_properties.as_ref().unwrap();
        assert!(
            std::str::from_utf8(shape_properties.as_xml())
                .unwrap()
                .contains("123456")
        );
        let extension_list = chart.extension_list.as_ref().unwrap();
        assert!(
            std::str::from_utf8(extension_list.as_xml())
                .unwrap()
                .contains(r#"xmlns:x="urn:example:chart-extension""#)
        );
        assert!(
            std::str::from_utf8(chart.plot_area.shape_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("654321")
        );
        assert!(
            std::str::from_utf8(chart.chart_extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("chartPayload")
        );

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        assert_eq!(reparsed.shape_properties, chart.shape_properties);
        assert_eq!(reparsed.text_properties, chart.text_properties);
        assert_eq!(reparsed.extension_list, chart.extension_list);
        assert_eq!(
            reparsed.plot_area.shape_properties,
            chart.plot_area.shape_properties
        );
        assert_eq!(
            reparsed.plot_area.extension_list,
            chart.plot_area.extension_list
        );
        assert_eq!(reparsed.chart_extension_list, chart.chart_extension_list);

        assert!(
            ChartShapeProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#
                    .to_vec()
            )
            .is_err()
        );
        assert!(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/><c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#
                    .to_vec()
            )
            .is_err()
        );
    }

    #[test]
    fn round_trips_chart_surface_shape_and_picture_options() {
        let xml = br#"<c:chartSpace
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:x="urn:example:surface">
            <c:chart><c:floor>
                <c:thickness val="64"/>
                <c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></c:spPr>
                <c:pictureOptions>
                    <c:applyToFront val="0"/><c:applyToSides/><c:applyToEnd val="1"/>
                    <c:pictureFormat val="stackScale"/><c:pictureStackUnit val="2.5"/>
                </c:pictureOptions>
                <c:extLst><c:ext uri="surface"><x:payload/></c:ext></c:extLst>
            </c:floor><c:plotArea/></c:chart>
        </c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let floor = chart.floor.as_ref().unwrap();
        assert_eq!(floor.thickness, Some(64));
        assert!(
            std::str::from_utf8(floor.shape_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("ABCDEF")
        );
        let options = floor.picture_options.as_ref().unwrap();
        assert_eq!(options.apply_to_front, Some(false));
        assert_eq!(options.apply_to_sides, Some(true));
        assert_eq!(options.apply_to_end, Some(true));
        assert_eq!(options.picture_format, Some(PictureFormat::StackScale));
        assert_eq!(options.picture_stack_unit, Some(2.5));
        assert!(
            std::str::from_utf8(floor.extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("urn:example:surface")
        );

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        let reparsed_floor = reparsed.floor.unwrap();
        assert_eq!(reparsed_floor.thickness, floor.thickness);
        assert_eq!(reparsed_floor.shape_properties, floor.shape_properties);
        assert_eq!(reparsed_floor.picture_options, floor.picture_options);
        assert_eq!(reparsed_floor.extension_list, floor.extension_list);

        let invalid = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:floor><c:pictureOptions><c:pictureFormat val="tile"/></c:pictureOptions></c:floor><c:plotArea/></c:chart></c:chartSpace>"#;
        assert!(parse_chart(invalid.as_slice()).is_err());

        let empty = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:view3D/><c:floor/><c:backWall/><c:sideWall/><c:plotArea/></c:chart></c:chartSpace>"#;
        let empty_chart = parse_chart(empty.as_slice()).unwrap();
        assert!(empty_chart.view_3d.is_some());
        assert!(empty_chart.floor.is_some());
        assert!(empty_chart.back_wall.is_some());
        assert!(empty_chart.side_wall.is_some());
    }

    #[test]
    fn round_trips_series_and_data_point_formatting_fragments() {
        let xml = br#"<c:chartSpace
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:x="urn:example:series">
            <c:chart><c:plotArea><c:barChart>
                <c:barDir val="col"/><c:grouping val="clustered"/><c:ser>
                    <c:idx val="0"/><c:order val="0"/>
                    <c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr>
                    <c:invertIfNegative val="1"/>
                    <c:pictureOptions><c:applyToFront val="0"/>
                        <c:pictureFormat val="stack"/><c:pictureStackUnit val="-2"/>
                    </c:pictureOptions>
                    <c:dPt><c:idx val="2"/>
                        <c:spPr><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></c:spPr>
                        <c:pictureOptions><c:applyToSides/></c:pictureOptions>
                        <c:extLst><c:ext uri="point"><x:pointPayload/></c:ext></c:extLst>
                    </c:dPt>
                    <c:shape val="cylinder"/>
                    <c:extLst><c:ext uri="series"><x:seriesPayload/></c:ext></c:extLst>
                </c:ser><c:axId val="1"/><c:axId val="2"/>
            </c:barChart></c:plotArea></c:chart>
        </c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let TypeGroup::Bar(group) = &chart.plot_area.type_groups[0] else {
            panic!("expected a bar chart");
        };
        let series = &group.common.series[0];
        assert!(
            std::str::from_utf8(series.shape_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("112233")
        );
        assert_eq!(
            series.picture_options.as_ref().unwrap().picture_format,
            Some(PictureFormat::Stack)
        );
        assert_eq!(
            series.picture_options.as_ref().unwrap().picture_stack_unit,
            Some(-2.0)
        );
        assert_eq!(series.bar_shape, Some(BarShape::Cylinder));
        assert!(
            std::str::from_utf8(series.extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("seriesPayload")
        );
        let point = &series.data_points[0];
        assert_eq!(
            point.picture_options.as_ref().unwrap().apply_to_sides,
            Some(true)
        );
        assert!(
            std::str::from_utf8(point.extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("pointPayload")
        );

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        let TypeGroup::Bar(reparsed_group) = &reparsed.plot_area.type_groups[0] else {
            panic!("expected a bar chart");
        };
        let reparsed_series = &reparsed_group.common.series[0];
        assert_eq!(reparsed_series.shape_properties, series.shape_properties);
        assert_eq!(reparsed_series.picture_options, series.picture_options);
        assert_eq!(reparsed_series.bar_shape, series.bar_shape);
        assert_eq!(reparsed_series.extension_list, series.extension_list);
        assert_eq!(
            reparsed_series.data_points[0].shape_properties,
            point.shape_properties
        );
        assert_eq!(
            reparsed_series.data_points[0].picture_options,
            point.picture_options
        );
        assert_eq!(
            reparsed_series.data_points[0].extension_list,
            point.extension_list
        );

        let unsupported =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:pictureOptions/></c:ser></c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>"#;
        let unsupported = parse_chart(unsupported.as_slice()).unwrap();
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &unsupported).is_err());

        let unsupported_shape =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:shape val="cone"/></c:ser></c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>"#;
        let unsupported_shape = parse_chart(unsupported_shape.as_slice()).unwrap();
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &unsupported_shape).is_err());

        for invalid_shape in [
            br#"<c:shape val="sphere"/>"#.as_slice(),
            br#"<c:shape/><c:shape/>"#.as_slice(),
        ] {
            let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>"#.to_vec();
            document.extend_from_slice(invalid_shape);
            document
                .extend_from_slice(b"</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>");
            assert!(parse_chart(document.as_slice()).is_err());
        }

        for supported in [
            br#"<c:areaChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:pictureOptions/></c:ser></c:areaChart>"#.as_slice(),
            br#"<c:area3DChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:pictureOptions/></c:ser></c:area3DChart>"#.as_slice(),
            br#"<c:bubbleChart><c:ser><c:idx val="0"/><c:order val="0"/><c:invertIfNegative/></c:ser></c:bubbleChart>"#.as_slice(),
        ] {
            let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>"#.to_vec();
            document.extend_from_slice(supported);
            document.extend_from_slice(b"</c:plotArea></c:chart></c:chartSpace>");
            let supported = parse_chart(document.as_slice()).unwrap();
            crate::charts::writer::write_chart(&mut Vec::new(), &supported).unwrap();
        }
    }

    #[test]
    fn round_trips_and_validates_chart_print_settings() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea/></c:chart><c:printSettings>
                <c:headerFooter alignWithMargins="0" differentOddEven="1" differentFirst="true">
                    <c:oddHeader>&amp;LRevenue</c:oddHeader><c:oddFooter>&amp;P / &amp;N</c:oddFooter>
                    <c:evenHeader/><c:firstFooter><![CDATA[First & last]]></c:firstFooter>
                </c:headerFooter>
                <c:pageMargins l="0.2" r="0.3" t="0.4" b="0.5" header="0.1" footer="0.15"/>
                <c:pageSetup paperSize="9" firstPageNumber="4" orientation="landscape"
                    blackAndWhite="1" draft="true" useFirstPageNumber="1"
                    horizontalDpi="300" verticalDpi="1200" copies="2"/>
            </c:printSettings></c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let settings = chart.print_settings.as_ref().unwrap();
        let header_footer = settings.header_footer.as_ref().unwrap();
        assert!(!header_footer.align_with_margins);
        assert!(header_footer.different_odd_even);
        assert!(header_footer.different_first);
        assert_eq!(header_footer.odd_header.as_deref(), Some("&LRevenue"));
        assert_eq!(header_footer.even_header.as_deref(), Some(""));
        assert_eq!(header_footer.first_footer.as_deref(), Some("First & last"));
        let margins = settings.page_margins.as_ref().unwrap();
        assert_eq!(margins.left, 0.2);
        assert_eq!(margins.footer, 0.15);
        let setup = settings.page_setup.as_ref().unwrap();
        assert_eq!(setup.paper_size, 9);
        assert_eq!(setup.first_page_number, 4);
        assert_eq!(setup.orientation, ChartPageOrientation::Landscape);
        assert!(setup.black_and_white);
        assert!(setup.draft);
        assert!(setup.use_first_page_number);
        assert_eq!(setup.horizontal_dpi, 300);
        assert_eq!(setup.vertical_dpi, 1200);
        assert_eq!(setup.copies, 2);

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        assert_eq!(
            reparsed
                .print_settings
                .as_ref()
                .unwrap()
                .header_footer
                .as_ref()
                .unwrap()
                .odd_footer
                .as_deref(),
            Some("&P / &N")
        );

        let empty =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea/></c:chart><c:printSettings/></c:chartSpace>"#;
        let empty = parse_chart(empty.as_slice()).unwrap();
        let empty = empty.print_settings.unwrap();
        assert!(empty.header_footer.is_none());
        assert!(empty.page_margins.is_none());
        assert!(empty.page_setup.is_none());

        for invalid in [
            br#"<c:pageMargins l="0.2" r="0.3" t="0.4" b="0.5" header="0.1"/>"#.as_slice(),
            br#"<c:pageSetup orientation="diagonal"/>"#.as_slice(),
            br#"<c:pageSetup/><c:pageSetup/>"#.as_slice(),
            br#"<c:headerFooter><c:bogus/></c:headerFooter>"#.as_slice(),
        ] {
            let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea/></c:chart><c:printSettings>"#.to_vec();
            document.extend_from_slice(invalid);
            document.extend_from_slice(b"</c:printSettings></c:chartSpace>");
            assert!(parse_chart(document.as_slice()).is_err());
        }

        let mut invalid = Chart::new();
        let mut settings = ChartPrintSettings::new();
        settings.page_margins = Some(ChartPageMargins::new(f64::NAN, 0.3, 0.4, 0.5, 0.1, 0.15));
        invalid.print_settings = Some(settings);
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &invalid).is_err());
    }

    #[test]
    fn round_trips_and_validates_pivot_chart_formats() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:x="urn:example:pivot-format">
            <c:chart><c:pivotFmts>
                <c:pivotFmt><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></c:spPr>
                    <c:txPr><a:bodyPr rot="600000"/><a:lstStyle/><a:p/></c:txPr><c:marker>
                    <c:symbol val="diamond"/><c:size val="8"/><c:spPr><a:ln w="25400"/></c:spPr>
                    <c:extLst><c:ext uri="marker"><x:markerPayload/></c:ext></c:extLst>
                </c:marker><c:dLbl><c:idx val="2"/><c:showVal val="1"/></c:dLbl>
                    <c:extLst><c:ext uri="pivot"><x:payload/></c:ext></c:extLst></c:pivotFmt>
                <c:pivotFmt><c:idx val="7"/><c:marker/></c:pivotFmt>
            </c:pivotFmts><c:plotArea/></c:chart></c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let formats = chart.pivot_formats.as_ref().unwrap();
        assert_eq!(formats.len(), 2);
        assert_eq!(formats[0].index, 2);
        let marker = formats[0].marker.as_ref().unwrap();
        assert_eq!(marker.symbol, Some(MarkerStyle::Diamond));
        assert_eq!(marker.size, Some(8));
        assert!(
            std::str::from_utf8(marker.shape_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("25400")
        );
        assert!(marker.extension_list.is_some());
        assert!(formats[0].data_label.as_ref().unwrap().show_value);
        assert!(
            std::str::from_utf8(formats[0].shape_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("123456")
        );
        assert!(
            std::str::from_utf8(formats[0].text_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("600000")
        );
        assert!(
            std::str::from_utf8(formats[0].extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("payload")
        );
        assert!(formats[1].marker.is_some());

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        assert_eq!(reparsed.pivot_formats.as_ref().unwrap().len(), 2);
        let reparsed_format = &reparsed.pivot_formats.as_ref().unwrap()[0];
        assert_eq!(
            reparsed_format.shape_properties,
            formats[0].shape_properties
        );
        assert_eq!(reparsed_format.text_properties, formats[0].text_properties);
        assert_eq!(reparsed_format.extension_list, formats[0].extension_list);

        let empty =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:pivotFmts/><c:plotArea/></c:chart></c:chartSpace>"#;
        assert!(
            parse_chart(empty.as_slice())
                .unwrap()
                .pivot_formats
                .unwrap()
                .is_empty()
        );

        let duplicate =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:pivotFmts><c:pivotFmt><c:idx val="1"/></c:pivotFmt>
                <c:pivotFmt><c:idx val="1"/></c:pivotFmt></c:pivotFmts>
                <c:plotArea/></c:chart></c:chartSpace>"#;
        assert!(parse_chart(duplicate.as_slice()).is_err());

        let duplicate_shape = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:pivotFmts><c:pivotFmt><c:idx val="1"/><c:spPr/><c:spPr/></c:pivotFmt></c:pivotFmts><c:plotArea/></c:chart></c:chartSpace>"#;
        assert!(parse_chart(duplicate_shape.as_slice()).is_err());

        let mut invalid = Chart::new();
        invalid.pivot_formats = Some(vec![PivotFormat::new(3), PivotFormat::new(3)]);
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &invalid).is_err());

        let mut invalid = Chart::new();
        let mut format = PivotFormat::new(3);
        format.marker = Some(Marker {
            symbol: None,
            size: Some(73),
            ..Marker::default()
        });
        invalid.pivot_formats = Some(vec![format]);
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &invalid).is_err());
    }

    #[test]
    fn preserves_explicit_empty_series_and_point_markers() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:marker/><c:dPt><c:idx val="0"/><c:marker/></c:dPt>
            </c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let TypeGroup::Line(group) = &chart.plot_area.type_groups[0] else {
            panic!("expected line chart");
        };
        let series = &group.common.series[0];
        assert!(series.marker_present);
        assert!(series.data_points[0].marker_present);

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        let TypeGroup::Line(group) = &reparsed.plot_area.type_groups[0] else {
            panic!("expected line chart");
        };
        assert!(group.common.series[0].marker_present);
        assert!(group.common.series[0].data_points[0].marker_present);
    }

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
                <c:f>Sheet<x:payload><c:style val="47"/><c:masterClrMapping/>ignored</x:payload>1!$A$1</c:f>
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
    fn round_trips_display_units_label_formatting_and_extensions() {
        let xml = br#"<c:chartSpace
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:x="urn:example:display-units">
            <c:chart><c:plotArea><c:valAx><c:axId val="1"/><c:scaling/>
                <c:axPos val="l"/><c:crossAx val="2"/><c:dispUnits>
                    <c:builtInUnit val="millions"/><c:dispUnitsLbl>
                        <c:layout/><c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></c:spPr>
                        <c:txPr><a:bodyPr rot="600000"/><a:lstStyle/><a:p/></c:txPr>
                    </c:dispUnitsLbl>
                    <c:extLst><c:ext uri="display-units"><x:payload/></c:ext></c:extLst>
                </c:dispUnits>
            </c:valAx></c:plotArea></c:chart>
        </c:chartSpace>"#;

        let chart = parse_chart(xml.as_slice()).unwrap();
        let Axis::Value(axis) = &chart.plot_area.axes[0] else {
            panic!("expected value axis");
        };
        let display_units = axis.display_units.as_ref().unwrap();
        assert!(display_units.show_label);
        assert!(
            std::str::from_utf8(
                display_units
                    .label_shape_properties
                    .as_ref()
                    .unwrap()
                    .as_xml()
            )
            .unwrap()
            .contains("ABCDEF")
        );
        assert!(
            std::str::from_utf8(
                display_units
                    .label_text_properties
                    .as_ref()
                    .unwrap()
                    .as_xml()
            )
            .unwrap()
            .contains("600000")
        );
        assert!(
            std::str::from_utf8(display_units.extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("payload")
        );

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        let Axis::Value(reparsed_axis) = &reparsed.plot_area.axes[0] else {
            panic!("expected value axis");
        };
        let reparsed_units = reparsed_axis.display_units.as_ref().unwrap();
        assert_eq!(
            reparsed_units.label_shape_properties,
            display_units.label_shape_properties
        );
        assert_eq!(
            reparsed_units.label_text_properties,
            display_units.label_text_properties
        );
        assert_eq!(reparsed_units.extension_list, display_units.extension_list);

        let duplicate = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:valAx><c:axId val="1"/><c:scaling/><c:axPos val="l"/><c:crossAx val="2"/><c:dispUnits><c:builtInUnit val="millions"/><c:dispUnitsLbl><c:spPr/><c:spPr/></c:dispUnitsLbl></c:dispUnits></c:valAx></c:plotArea></c:chart></c:chartSpace>"#;
        assert!(parse_chart(duplicate.as_slice()).is_err());
    }

    #[test]
    fn parses_and_validates_chart_data_tables() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:x="urn:example:data-table">
            <c:chart><c:plotArea><c:lineChart></c:lineChart>
                <c:dTable><c:showHorzBorder/><c:showVertBorder val="0"/>
                    <c:showOutline val="true"/><c:showKeys val="false"/>
                    <c:spPr><a:solidFill><a:srgbClr val="F0E0D0"/></a:solidFill></c:spPr>
                    <c:txPr><a:bodyPr/><a:lstStyle/><a:p/></c:txPr>
                    <c:extLst><c:ext uri="table"><x:payload/></c:ext></c:extLst>
                </c:dTable>
            </c:plotArea></c:chart>
        </c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let table = chart.plot_area.data_table.as_ref().unwrap();
        assert!(table.show_horizontal_border);
        assert!(!table.show_vertical_border);
        assert!(table.show_outline);
        assert!(!table.show_legend_keys);
        assert!(
            std::str::from_utf8(table.shape_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("F0E0D0")
        );
        assert!(table.text_properties.is_some());
        assert!(
            std::str::from_utf8(table.extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("urn:example:data-table")
        );

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        let reparsed = reparsed.plot_area.data_table.unwrap();
        assert_eq!(reparsed.shape_properties, table.shape_properties);
        assert_eq!(reparsed.text_properties, table.text_properties);
        assert_eq!(reparsed.extension_list, table.extension_list);

        let duplicate =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:lineChart></c:lineChart><c:dTable>
                <c:showKeys/><c:showKeys val="0"/>
            </c:dTable></c:plotArea></c:chart>
        </c:chartSpace>"#;
        assert!(parse_chart(duplicate.as_slice()).is_err());
    }

    #[test]
    fn parses_and_validates_chart_group_data_labels() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea>
                <c:lineChart><c:dLbls><c:numFmt formatCode="0.0%" sourceLinked="0"/>
                    <c:dLblPos val="r"/><c:showVal/><c:showCatName val="true"/>
                    <c:separator> / </c:separator><c:showLeaderLines/>
                </c:dLbls></c:lineChart>
                <c:area3DChart><c:dLbls/></c:area3DChart>
            </c:plotArea></c:chart>
        </c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let TypeGroup::Line(line) = &chart.plot_area.type_groups[0] else {
            panic!("expected line chart");
        };
        let labels = line.common.data_labels.as_ref().unwrap();
        assert_eq!(labels.position, Some(DataLabelPosition::Right));
        assert!(labels.show_value);
        assert!(labels.show_category_name);
        assert!(labels.show_leader_lines);
        assert_eq!(labels.separator.as_deref(), Some(" / "));
        assert_eq!(labels.number_format.as_ref().unwrap().format_code, "0.0%");
        assert!(!labels.number_format.as_ref().unwrap().source_linked);
        let TypeGroup::Area3D(area) = &chart.plot_area.type_groups[1] else {
            panic!("expected 3D area chart");
        };
        assert!(area.common.data_labels.is_some());

        let duplicate =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:barChart><c:dLbls/><c:dLbls/>
            </c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        assert!(parse_chart(duplicate.as_slice()).is_err());
    }

    #[test]
    fn preserves_and_validates_chart_group_axis_bindings() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea>
                <c:lineChart><c:axId val="17"/><c:axId val="29"/></c:lineChart>
                <c:area3DChart><c:axId val="31"/><c:axId val="37"/><c:axId val="41"/></c:area3DChart>
            </c:plotArea></c:chart>
        </c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let TypeGroup::Line(line) = &chart.plot_area.type_groups[0] else {
            panic!("expected line chart");
        };
        assert_eq!(line.common.axis_ids, [17, 29]);
        let TypeGroup::Area3D(area) = &chart.plot_area.type_groups[1] else {
            panic!("expected 3D area chart");
        };
        assert_eq!(area.common.axis_ids, [31, 37, 41]);

        for axis_ids in [vec![7], vec![7, 7], vec![1, 2, 3]] {
            let mut chart = Chart::new();
            let mut line = LineTypeGroup::new(BarGrouping::Standard);
            line.common.axis_ids = axis_ids;
            chart.plot_area.type_groups.push(TypeGroup::Line(line));
            assert!(crate::charts::writer::write_chart(&mut Vec::new(), &chart).is_err());
        }
    }

    #[test]
    fn round_trips_chart_lines_and_up_down_bars() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:x="urn:example:up-down-bars">
            <c:chart><c:plotArea>
                <c:areaChart><c:dropLines/></c:areaChart>
                <c:area3DChart><c:dropLines/></c:area3DChart>
                <c:barChart><c:serLines><c:spPr><a:solidFill><a:srgbClr val="333333"/></a:solidFill></c:spPr></c:serLines><c:serLines/></c:barChart>
                <c:lineChart><c:dropLines><c:spPr><a:solidFill><a:srgbClr val="111111"/></a:solidFill></c:spPr></c:dropLines><c:hiLowLines/><c:upDownBars>
                    <c:gapWidth val="225%"/><c:upBars><c:spPr><a:solidFill><a:srgbClr val="222222"/></a:solidFill></c:spPr></c:upBars><c:downBars/>
                    <c:extLst><c:ext uri="bars"><x:payload/></c:ext></c:extLst>
                </c:upDownBars></c:lineChart>
                <c:line3DChart><c:dropLines/></c:line3DChart>
                <c:ofPieChart><c:ofPieType val="bar"/><c:serLines/></c:ofPieChart>
                <c:stockChart><c:dropLines/><c:hiLowLines/><c:upDownBars/></c:stockChart>
            </c:plotArea></c:chart>
        </c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let TypeGroup::Area(area) = &chart.plot_area.type_groups[0] else {
            panic!("expected area chart");
        };
        assert!(area.drop_lines.is_some());
        let TypeGroup::Area3D(area) = &chart.plot_area.type_groups[1] else {
            panic!("expected 3D area chart");
        };
        assert!(area.drop_lines.is_some());
        let TypeGroup::Bar(bar) = &chart.plot_area.type_groups[2] else {
            panic!("expected bar chart");
        };
        assert_eq!(bar.series_lines.len(), 2);
        assert!(
            std::str::from_utf8(
                bar.series_lines[0]
                    .shape_properties
                    .as_ref()
                    .unwrap()
                    .as_xml()
            )
            .unwrap()
            .contains("333333")
        );
        let TypeGroup::Line(line) = &chart.plot_area.type_groups[3] else {
            panic!("expected line chart");
        };
        assert!(line.drop_lines.is_some());
        assert!(line.high_low_lines.is_some());
        assert!(
            std::str::from_utf8(
                line.drop_lines
                    .as_ref()
                    .unwrap()
                    .shape_properties
                    .as_ref()
                    .unwrap()
                    .as_xml()
            )
            .unwrap()
            .contains("111111")
        );
        let bars = line.up_down_bars.as_ref().unwrap();
        assert_eq!(bars.gap_width, Some(225));
        assert!(bars.up_bars.is_some());
        assert!(bars.down_bars.is_some());
        assert!(
            std::str::from_utf8(
                bars.up_bars
                    .as_ref()
                    .unwrap()
                    .shape_properties
                    .as_ref()
                    .unwrap()
                    .as_xml()
            )
            .unwrap()
            .contains("222222")
        );
        assert!(
            std::str::from_utf8(bars.extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("urn:example:up-down-bars")
        );
        let TypeGroup::Line3D(line) = &chart.plot_area.type_groups[4] else {
            panic!("expected 3D line chart");
        };
        assert!(line.drop_lines.is_some());
        let TypeGroup::OfPie(of_pie) = &chart.plot_area.type_groups[5] else {
            panic!("expected of-pie chart");
        };
        assert_eq!(of_pie.series_lines.len(), 1);
        let TypeGroup::Stock(stock) = &chart.plot_area.type_groups[6] else {
            panic!("expected stock chart");
        };
        assert!(stock.drop_lines.is_some());
        assert!(stock.high_low_lines.is_some());
        assert!(stock.up_down_bars.is_some());

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        let TypeGroup::Line(line) = &reparsed.plot_area.type_groups[3] else {
            panic!("expected line chart");
        };
        assert_eq!(line.up_down_bars.as_ref().unwrap().gap_width, Some(225));
        assert_eq!(
            line.up_down_bars.as_ref().unwrap().extension_list,
            bars.extension_list
        );
        let TypeGroup::Bar(bar) = &reparsed.plot_area.type_groups[2] else {
            panic!("expected bar chart");
        };
        assert_eq!(bar.series_lines.len(), 2);
        assert!(
            std::str::from_utf8(
                bar.series_lines[0]
                    .shape_properties
                    .as_ref()
                    .unwrap()
                    .as_xml()
            )
            .unwrap()
            .contains("333333")
        );

        let duplicate =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:lineChart><c:dropLines/><c:dropLines/>
            </c:lineChart></c:plotArea></c:chart></c:chartSpace>"#;
        assert!(parse_chart(duplicate.as_slice()).is_err());

        let duplicate_formatting =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:lineChart><c:dropLines><c:spPr/><c:spPr/></c:dropLines>
            </c:lineChart></c:plotArea></c:chart></c:chartSpace>"#;
        assert!(parse_chart(duplicate_formatting.as_slice()).is_err());

        let mut invalid = Chart::new();
        let mut line = LineTypeGroup::new(BarGrouping::Standard);
        line.up_down_bars = Some(UpDownBars {
            gap_width: Some(501),
            ..UpDownBars::default()
        });
        invalid.plot_area.type_groups.push(TypeGroup::Line(line));
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &invalid).is_err());
    }

    #[test]
    fn round_trips_and_validates_surface_band_formats() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea>
                <c:surfaceChart><c:wireframe/><c:bandFmts>
                    <c:bandFmt><c:idx val="2"/><c:spPr/></c:bandFmt>
                    <c:bandFmt><c:idx val="7"/><c:spPr/></c:bandFmt>
                </c:bandFmts></c:surfaceChart>
                <c:surface3DChart><c:bandFmts/></c:surface3DChart>
            </c:plotArea></c:chart>
        </c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let TypeGroup::Surface(surface) = &chart.plot_area.type_groups[0] else {
            panic!("expected surface chart");
        };
        assert!(surface.wireframe);
        assert_eq!(
            surface
                .band_formats
                .as_ref()
                .unwrap()
                .iter()
                .map(|format| format.index)
                .collect::<Vec<_>>(),
            [2, 7]
        );
        let TypeGroup::Surface3D(surface) = &chart.plot_area.type_groups[1] else {
            panic!("expected 3D surface chart");
        };
        assert!(surface.band_formats.as_ref().unwrap().is_empty());

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        let TypeGroup::Surface(surface) = &reparsed.plot_area.type_groups[0] else {
            panic!("expected surface chart");
        };
        assert_eq!(surface.band_formats.as_ref().unwrap().len(), 2);

        let duplicate =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:surfaceChart><c:bandFmts>
                <c:bandFmt><c:idx val="1"/></c:bandFmt>
                <c:bandFmt><c:idx val="1"/></c:bandFmt>
            </c:bandFmts></c:surfaceChart></c:plotArea></c:chart></c:chartSpace>"#;
        assert!(parse_chart(duplicate.as_slice()).is_err());

        let mut invalid = Chart::new();
        let mut surface = SurfaceTypeGroup::new();
        surface.band_formats = Some(vec![BandFormat::new(3), BandFormat::new(3)]);
        invalid
            .plot_area
            .type_groups
            .push(TypeGroup::Surface(surface));
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &invalid).is_err());
    }

    #[test]
    fn parses_of_pie_schema_defaults_and_empty_custom_split() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:ofPieChart><c:ofPieType/>
                <c:gapWidth/><c:splitType/><c:custSplit/><c:secondPieSize/>
            </c:ofPieChart></c:plotArea></c:chart>
        </c:chartSpace>"#;

        let chart = parse_chart(xml.as_slice()).unwrap();
        let [TypeGroup::OfPie(group)] = chart.plot_area.type_groups.as_slice() else {
            panic!("expected an of-pie chart");
        };
        assert_eq!(group.of_pie_type, OfPieType::Pie);
        assert_eq!(group.gap_width, Some(150));
        assert_eq!(group.split_type, Some(OfPieSplitType::Automatic));
        assert_eq!(group.custom_split_points.as_deref(), Some([].as_slice()));
        assert_eq!(group.second_pie_size, Some(75));

        let percent_xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:ofPieChart><c:ofPieType val="bar"/>
                <c:gapWidth val="225%"/><c:secondPieSize val="80%"/>
            </c:ofPieChart></c:plotArea></c:chart>
        </c:chartSpace>"#;
        let chart = parse_chart(percent_xml.as_slice()).unwrap();
        let [TypeGroup::OfPie(group)] = chart.plot_area.type_groups.as_slice() else {
            panic!("expected an of-pie chart");
        };
        assert_eq!(group.gap_width, Some(225));
        assert_eq!(group.second_pie_size, Some(80));
    }

    #[test]
    fn rejects_invalid_of_pie_input_and_output() {
        for xml in [
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:ofPieChart/></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:ofPieChart><c:ofPieType val="ring"/></c:ofPieChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:ofPieChart><c:ofPieType/><c:gapWidth val="501"/></c:ofPieChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:ofPieChart><c:ofPieType/><c:secondPieSize val="4"/></c:ofPieChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:ofPieChart><c:ofPieType/><c:splitPos val="NaN"/></c:ofPieChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
        ] {
            assert!(parse_chart(xml).is_err());
        }

        for invalid_group in [
            OfPieTypeGroup {
                gap_width: Some(501),
                ..OfPieTypeGroup::new(OfPieType::Pie)
            },
            OfPieTypeGroup {
                split_position: Some(f64::INFINITY),
                ..OfPieTypeGroup::new(OfPieType::Pie)
            },
            OfPieTypeGroup {
                second_pie_size: Some(4),
                ..OfPieTypeGroup::new(OfPieType::Pie)
            },
        ] {
            let mut chart = Chart::new();
            chart
                .plot_area
                .type_groups
                .push(TypeGroup::OfPie(invalid_group));
            assert!(crate::charts::writer::write_chart(&mut Vec::new(), &chart).is_err());
        }
    }

    #[test]
    fn parses_chart_group_percentage_union_values() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea>
                <c:barChart><c:barDir val="col"/><c:gapWidth val="225%"/><c:overlap val="-25%"/></c:barChart>
                <c:bar3DChart><c:barDir val="bar"/><c:gapWidth/><c:gapDepth val="175%"/><c:shape/></c:bar3DChart>
                <c:bubbleChart><c:bubbleScale val="125%"/><c:sizeRepresents/></c:bubbleChart>
                <c:doughnutChart><c:firstSliceAng/><c:holeSize val="5%"/></c:doughnutChart>
                <c:area3DChart><c:gapDepth val="500%"/></c:area3DChart>
                <c:lineChart><c:smooth/></c:lineChart>
                <c:line3DChart><c:gapDepth/></c:line3DChart>
            </c:plotArea></c:chart>
        </c:chartSpace>"#;

        let chart = parse_chart(xml.as_slice()).unwrap();
        let TypeGroup::Bar(bar) = &chart.plot_area.type_groups[0] else {
            panic!("expected bar chart");
        };
        assert_eq!(bar.gap_width, Some(225));
        assert_eq!(bar.overlap, Some(-25));
        let TypeGroup::Bar3D(bar) = &chart.plot_area.type_groups[1] else {
            panic!("expected 3D bar chart");
        };
        assert_eq!(bar.gap_width, Some(150));
        assert_eq!(bar.gap_depth, Some(175));
        assert_eq!(bar.shape, Some(BarShape::Box));
        let TypeGroup::Bubble(bubble) = &chart.plot_area.type_groups[2] else {
            panic!("expected bubble chart");
        };
        assert_eq!(bubble.bubble_scale, Some(125));
        assert_eq!(bubble.size_represents, "area");
        let TypeGroup::Doughnut(doughnut) = &chart.plot_area.type_groups[3] else {
            panic!("expected doughnut chart");
        };
        assert_eq!(doughnut.first_slice_angle, 0);
        assert_eq!(doughnut.hole_size, 5);
        let TypeGroup::Area3D(area) = &chart.plot_area.type_groups[4] else {
            panic!("expected 3D area chart");
        };
        assert_eq!(area.gap_depth, Some(500));
        let TypeGroup::Line(line) = &chart.plot_area.type_groups[5] else {
            panic!("expected line chart");
        };
        assert!(line.smooth);
        let TypeGroup::Line3D(line) = &chart.plot_area.type_groups[6] else {
            panic!("expected 3D line chart");
        };
        assert_eq!(line.gap_depth, Some(150));
    }

    #[test]
    fn writer_rejects_invalid_chart_group_ranges() {
        let mut bar = BarTypeGroup::new(BarDirection::Column, BarGrouping::Clustered);
        bar.gap_width = Some(501);
        let mut bar_3d = Bar3DTypeGroup::new(BarDirection::Column, BarGrouping::Clustered);
        bar_3d.gap_depth = Some(501);
        let mut bubble = BubbleTypeGroup::new();
        bubble.size_represents = "diameter".to_string();
        let mut doughnut = DoughnutTypeGroup::new();
        doughnut.hole_size = 0;
        let mut line_3d = Line3DTypeGroup::new(BarGrouping::Standard);
        line_3d.gap_depth = Some(501);
        let mut pie = PieTypeGroup::new();
        pie.first_slice_angle = 361;

        for group in [
            TypeGroup::Bar(bar),
            TypeGroup::Bar3D(bar_3d),
            TypeGroup::Bubble(bubble),
            TypeGroup::Doughnut(doughnut),
            TypeGroup::Line3D(line_3d),
            TypeGroup::Pie(pie),
        ] {
            let mut chart = Chart::new();
            chart.plot_area.type_groups.push(group);
            assert!(crate::charts::writer::write_chart(&mut Vec::new(), &chart).is_err());
        }
    }

    #[test]
    fn writer_rejects_invalid_axes_and_duplicate_legend_entries() {
        let mut chart = Chart::new();
        let mut axis = ValueAxis::new(1, AxisPosition::Left, 2);
        let mut units = DisplayUnits::custom(1_000.0);
        units.built_in_unit = Some(BuiltInUnit::Thousands);
        axis.display_units = Some(Box::new(units));
        chart.plot_area.axes.push(Axis::Value(axis));
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &chart).is_err());

        let mut chart = Chart::new();
        let mut axis = CategoryAxis::new(1, AxisPosition::Bottom, 2);
        axis.min = Some(2.0);
        axis.max = Some(1.0);
        chart.plot_area.axes.push(Axis::Category(axis));
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &chart).is_err());

        let mut chart = Chart::new();
        let mut axis = ValueAxis::new(1, AxisPosition::Left, 2);
        axis.display_units = Some(Box::new(DisplayUnits::custom(f64::NAN)));
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
    fn round_trips_legend_and_entry_formatting_fragments() {
        let xml = br#"<c:chartSpace
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:x="urn:example:legend">
            <c:chart><c:plotArea/><c:legend><c:legendPos val="b"/>
                <c:legendEntry><c:idx val="2"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p/></c:txPr>
                    <c:extLst><c:ext uri="entry"><x:entryPayload/></c:ext></c:extLst>
                </c:legendEntry>
                <c:legendEntry><c:idx val="3"/><c:delete val="1"/></c:legendEntry>
                <c:overlay val="1"/>
                <c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></c:spPr>
                <c:txPr><a:bodyPr rot="1200000"/><a:lstStyle/><a:p/></c:txPr>
                <c:extLst><c:ext uri="legend"><x:legendPayload/></c:ext></c:extLst>
            </c:legend></c:chart>
        </c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        let legend = chart.legend.as_ref().unwrap();
        assert_eq!(legend.position, LegendPosition::Bottom);
        assert!(legend.overlay);
        assert_eq!(legend.entries.len(), 2);
        assert!(legend.entries[0].text_properties.is_some());
        assert!(legend.entries[0].extension_list.is_some());
        assert!(legend.entries[1].deleted);
        assert!(
            std::str::from_utf8(legend.shape_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("123456")
        );
        assert!(
            std::str::from_utf8(legend.text_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("1200000")
        );
        assert!(
            std::str::from_utf8(legend.extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("legendPayload")
        );

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        let reparsed = reparsed.legend.as_ref().unwrap();
        assert_eq!(reparsed.shape_properties, legend.shape_properties);
        assert_eq!(reparsed.text_properties, legend.text_properties);
        assert_eq!(reparsed.extension_list, legend.extension_list);
        assert_eq!(
            reparsed.entries[0].text_properties,
            legend.entries[0].text_properties
        );
        assert_eq!(
            reparsed.entries[0].extension_list,
            legend.entries[0].extension_list
        );

        for invalid_entry in [
            br#"<c:legendEntry><c:idx val="1"/></c:legendEntry>"#.as_slice(),
            br#"<c:legendEntry><c:idx val="1"/><c:delete/><c:txPr/></c:legendEntry>"#.as_slice(),
        ] {
            let mut document = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea/><c:legend>"#.to_vec();
            document.extend_from_slice(invalid_entry);
            document.extend_from_slice(b"</c:legend></c:chart></c:chartSpace>");
            assert!(parse_chart(document.as_slice()).is_err());
        }

        let mut invalid = Chart::new();
        let mut legend = Legend::default();
        let mut entry = LegendEntry::new(1);
        entry.deleted = true;
        entry.text_properties = Some(
            ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#
                    .to_vec(),
            )
            .unwrap(),
        );
        legend.entries.push(entry);
        invalid.legend = Some(legend);
        assert!(crate::charts::writer::write_chart(&mut Vec::new(), &invalid).is_err());
    }

    #[test]
    fn round_trips_chart_and_axis_title_formatting_fragments() {
        let xml = br#"<c:chartSpace
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:x="urn:example:title">
            <c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx>
                <c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr>
                <c:txPr><a:bodyPr rot="600000"/><a:lstStyle/><a:p/></c:txPr>
                <c:extLst><c:ext uri="chart-title"><x:chartPayload/></c:ext></c:extLst>
            </c:title><c:plotArea><c:catAx><c:axId val="1"/><c:scaling>
                <c:extLst><c:ext uri="scaling"><x:scalingPayload/></c:ext></c:extLst>
            </c:scaling>
                <c:axPos val="b"/><c:majorGridlines><c:spPr><a:solidFill><a:srgbClr val="445566"/></a:solidFill></c:spPr></c:majorGridlines>
                <c:minorGridlines/><c:title><c:tx><c:rich><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx>
                    <c:spPr><a:noFill/></c:spPr><c:txPr><a:bodyPr vert="vert"/><a:lstStyle/><a:p/></c:txPr>
                    <c:extLst><c:ext uri="axis-title"><x:axisPayload/></c:ext></c:extLst>
                </c:title><c:spPr><a:ln w="12700"/></c:spPr>
                <c:txPr><a:bodyPr rot="-600000"/><a:lstStyle/><a:p/></c:txPr>
                <c:crossAx val="2"/><c:extLst><c:ext uri="axis"><x:axisBodyPayload/></c:ext></c:extLst>
                </c:catAx></c:plotArea></c:chart>
        </c:chartSpace>"#;
        let chart = parse_chart(xml.as_slice()).unwrap();
        assert!(
            std::str::from_utf8(chart.title_shape_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("112233")
        );
        assert!(
            std::str::from_utf8(chart.title_text_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("600000")
        );
        assert!(
            std::str::from_utf8(chart.title_extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("chartPayload")
        );
        let common = chart.plot_area.axes[0].common();
        assert!(common.show_major_gridlines);
        assert!(common.show_minor_gridlines);
        assert!(
            std::str::from_utf8(
                common
                    .major_gridlines
                    .as_ref()
                    .unwrap()
                    .shape_properties
                    .as_ref()
                    .unwrap()
                    .as_xml()
            )
            .unwrap()
            .contains("445566")
        );
        assert!(common.minor_gridlines.is_some());
        assert!(
            std::str::from_utf8(common.shape_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("12700")
        );
        assert!(
            std::str::from_utf8(common.text_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("-600000")
        );
        assert!(
            std::str::from_utf8(common.scaling_extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("scalingPayload")
        );
        assert!(
            std::str::from_utf8(common.extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("axisBodyPayload")
        );
        assert!(common.title_shape_properties.is_some());
        assert!(
            std::str::from_utf8(common.title_text_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("vert")
        );
        assert!(
            std::str::from_utf8(common.title_extension_list.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("axisPayload")
        );

        let mut output = Vec::new();
        crate::charts::writer::write_chart(&mut output, &chart).unwrap();
        let reparsed = parse_chart(output.as_slice()).unwrap();
        assert_eq!(
            reparsed.title_shape_properties,
            chart.title_shape_properties
        );
        assert_eq!(reparsed.title_text_properties, chart.title_text_properties);
        assert_eq!(reparsed.title_extension_list, chart.title_extension_list);
        let reparsed_common = reparsed.plot_area.axes[0].common();
        assert_eq!(
            reparsed_common.title_shape_properties,
            common.title_shape_properties
        );
        assert_eq!(
            reparsed_common.title_text_properties,
            common.title_text_properties
        );
        assert_eq!(
            reparsed_common.title_extension_list,
            common.title_extension_list
        );
        assert_eq!(reparsed_common.major_gridlines, common.major_gridlines);
        assert_eq!(reparsed_common.minor_gridlines, common.minor_gridlines);
        assert_eq!(reparsed_common.shape_properties, common.shape_properties);
        assert_eq!(reparsed_common.text_properties, common.text_properties);
        assert_eq!(
            reparsed_common.scaling_extension_list,
            common.scaling_extension_list
        );
        assert_eq!(reparsed_common.extension_list, common.extension_list);

        let duplicate = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:title><c:spPr/><c:spPr/></c:title><c:plotArea/></c:chart></c:chartSpace>"#;
        assert!(parse_chart(duplicate.as_slice()).is_err());

        let duplicate_gridlines = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:majorGridlines/><c:majorGridlines/><c:crossAx val="2"/></c:catAx></c:plotArea></c:chart></c:chartSpace>"#;
        assert!(parse_chart(duplicate_gridlines.as_slice()).is_err());
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
    fn rejects_empty_plot_containers_without_consuming_following_content() {
        for xml in [
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart><c:plotArea><c:lineChart/><c:dTable/></c:plotArea></c:chart>
            </c:chartSpace>"#
                .as_slice(),
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart><c:plotArea><c:catAx/><c:dTable/></c:plotArea></c:chart>
            </c:chartSpace>"#
                .as_slice(),
        ] {
            assert!(parse_chart(xml).is_err());
        }

        let empty_layout =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart><c:plotArea><c:layout/><c:lineChart></c:lineChart>
                </c:plotArea></c:chart>
            </c:chartSpace>"#;
        let chart = parse_chart(empty_layout.as_slice()).unwrap();
        assert!(chart.plot_area.layout.is_some());
    }

    #[test]
    fn writer_round_trips_every_modeled_chart_group() {
        let mut area_3d = Area3DTypeGroup::new(BarGrouping::Stacked);
        area_3d.gap_depth = Some(175);
        area_3d.common.axis_ids = vec![10, 20, 30];
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
        scatter_series.marker_shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        scatter_series.marker_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:marker"><c:ext uri="series-marker"><x:payload/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        scatter_series.smooth = true;
        let mut point = DataPoint::new(2).with_marker(7, MarkerStyle::Diamond);
        point.marker_shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ln w="12700"/></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        point.marker_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:marker"><c:ext uri="point-marker"><x:payload/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
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
        labels.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        labels.text_properties = Some(
            ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr rot="1200000"/><a:lstStyle/><a:p/></c:txPr>"#.to_vec(),
            )
            .unwrap(),
        );
        labels.leader_lines = Some(ChartLines {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ln w="38100"/></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
        });
        labels.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:labels"><c:ext uri="labels"><x:payload/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        labels.separator = Some(" & ".to_string());
        let mut point_label = DataLabel::new(2);
        point_label.layout = Some(Layout::new().with_position(0.6, 0.7));
        point_label.text = Some(TitleText::from_ref("Sheet1!$E$2"));
        point_label.number_format = Some(NumberFormat::new("$0.00"));
        point_label.position = Some(DataLabelPosition::Left);
        point_label.show_category_name = true;
        point_label.separator = Some(" / ".to_string());
        point_label.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:noFill/></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        point_label.text_properties = Some(
            ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr vert="vert"/><a:lstStyle/><a:p/></c:txPr>"#.to_vec(),
            )
            .unwrap(),
        );
        point_label.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:labels"><c:ext uri="point-label"><x:payload/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
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
        let mut of_pie = OfPieTypeGroup::new(OfPieType::Bar);
        of_pie.common.vary_colors = true;
        of_pie.gap_width = Some(225);
        of_pie.split_type = Some(OfPieSplitType::Custom);
        of_pie.split_position = Some(3.5);
        of_pie.custom_split_points = Some(vec![1, 4]);
        of_pie.second_pie_size = Some(80);
        let mut line = LineTypeGroup::new(BarGrouping::Standard);
        line.smooth = true;
        line.common.axis_ids = vec![41, 42];
        let mut group_labels = DataLabels::new()
            .with_position(DataLabelPosition::Right)
            .with_show_value(true);
        group_labels.show_category_name = true;
        group_labels.separator = Some(" | ".to_string());
        line.common.data_labels = Some(group_labels);
        let mut line_3d = Line3DTypeGroup::new(BarGrouping::PercentStacked);
        line_3d.gap_depth = Some(210);
        line_3d.common.axis_ids = vec![50, 51, 52];

        let mut chart = Chart::new();
        chart.plot_area.data_table = Some(DataTable {
            show_horizontal_border: true,
            show_vertical_border: false,
            show_outline: true,
            show_legend_keys: true,
            ..DataTable::default()
        });
        chart.plot_area.type_groups = vec![
            TypeGroup::Area(AreaTypeGroup::new(BarGrouping::Standard)),
            TypeGroup::Area3D(area_3d),
            TypeGroup::Bar(BarTypeGroup::new(
                BarDirection::Column,
                BarGrouping::Clustered,
            )),
            TypeGroup::Bar3D(Bar3DTypeGroup::new(BarDirection::Bar, BarGrouping::Stacked)),
            TypeGroup::Bubble(bubble),
            TypeGroup::Doughnut(doughnut),
            TypeGroup::Line(line),
            TypeGroup::Line3D(line_3d),
            TypeGroup::OfPie(of_pie),
            TypeGroup::Pie(PieTypeGroup::new()),
            TypeGroup::Pie3D(Pie3DTypeGroup::new()),
            TypeGroup::Radar(RadarTypeGroup::new(RadarStyle::Filled)),
            TypeGroup::Scatter(scatter),
            TypeGroup::Stock(StockTypeGroup::new()),
            TypeGroup::Surface(surface),
            TypeGroup::Surface3D(surface_3d),
        ];
        for (index, group) in chart.plot_area.type_groups.iter_mut().enumerate() {
            group.common_mut().extension_list = Some(
                ChartExtensionList::from_xml(
                    format!(
                        r#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:example:group"><c:ext uri="group-{index}"><x:payload/></c:ext></c:extLst>"#
                    )
                    .into_bytes(),
                )
                .unwrap(),
            );
        }

        let mut xml = Vec::new();
        crate::charts::writer::write_chart(&mut xml, &chart).unwrap();
        let parsed = parse_chart(xml.as_slice()).unwrap();

        assert_eq!(parsed.plot_area.type_groups.len(), 16);
        for (index, group) in parsed.plot_area.type_groups.iter().enumerate() {
            assert!(
                std::str::from_utf8(group.common().extension_list.as_ref().unwrap().as_xml())
                    .unwrap()
                    .contains(&format!(r#"uri="group-{index}""#))
            );
        }
        let data_table = parsed.plot_area.data_table.as_ref().unwrap();
        assert!(data_table.show_horizontal_border);
        assert!(!data_table.show_vertical_border);
        assert!(data_table.show_outline);
        assert!(data_table.show_legend_keys);
        assert!(matches!(
            parsed.plot_area.type_groups[0],
            TypeGroup::Area(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[1],
            TypeGroup::Area3D(_)
        ));
        let TypeGroup::Area3D(group) = &parsed.plot_area.type_groups[1] else {
            unreachable!();
        };
        assert_eq!(group.gap_depth, Some(175));
        assert_eq!(group.common.axis_ids, [10, 20, 30]);
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
        let TypeGroup::Line(group) = &parsed.plot_area.type_groups[6] else {
            unreachable!();
        };
        assert!(group.smooth);
        assert_eq!(group.common.axis_ids, [41, 42]);
        let labels = group.common.data_labels.as_ref().unwrap();
        assert_eq!(labels.position, Some(DataLabelPosition::Right));
        assert!(labels.show_value);
        assert!(labels.show_category_name);
        assert_eq!(labels.separator.as_deref(), Some(" | "));
        let TypeGroup::Line3D(group) = &parsed.plot_area.type_groups[7] else {
            unreachable!();
        };
        assert_eq!(group.gap_depth, Some(210));
        assert_eq!(group.common.axis_ids, [50, 51, 52]);
        assert!(matches!(parsed.plot_area.type_groups[9], TypeGroup::Pie(_)));
        assert!(matches!(
            parsed.plot_area.type_groups[10],
            TypeGroup::Pie3D(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[11],
            TypeGroup::Radar(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[13],
            TypeGroup::Stock(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[14],
            TypeGroup::Surface(_)
        ));
        assert!(matches!(
            parsed.plot_area.type_groups[15],
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
        let TypeGroup::OfPie(group) = &parsed.plot_area.type_groups[8] else {
            unreachable!();
        };
        assert_eq!(group.of_pie_type, OfPieType::Bar);
        assert!(group.common.vary_colors);
        assert_eq!(group.gap_width, Some(225));
        assert_eq!(group.split_type, Some(OfPieSplitType::Custom));
        assert_eq!(group.split_position, Some(3.5));
        assert_eq!(group.custom_split_points.as_deref(), Some(&[1, 4][..]));
        assert_eq!(group.second_pie_size, Some(80));
        let TypeGroup::Scatter(group) = &parsed.plot_area.type_groups[12] else {
            unreachable!();
        };
        let series = &group.common.series[0];
        assert_eq!(series.marker_symbol, Some(MarkerStyle::Star));
        assert_eq!(series.marker_size, Some(9));
        assert!(
            std::str::from_utf8(series.marker_shape_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("AABBCC")
        );
        assert!(series.marker_extension_list.is_some());
        assert!(series.smooth);
        assert_eq!(series.data_points.len(), 1);
        assert_eq!(series.data_points[0].index, 2);
        assert_eq!(series.data_points[0].marker_size, Some(7));
        assert_eq!(
            series.data_points[0].marker_symbol,
            Some(MarkerStyle::Diamond)
        );
        assert!(
            std::str::from_utf8(
                series.data_points[0]
                    .marker_shape_properties
                    .as_ref()
                    .unwrap()
                    .as_xml()
            )
            .unwrap()
            .contains("12700")
        );
        assert!(series.data_points[0].marker_extension_list.is_some());
        assert!(series.data_points[0].invert_if_negative);
        assert_eq!(series.data_points[0].bubble_3d, Some(false));
        assert_eq!(series.data_points[0].explosion, Some(15));
        let labels = series.data_labels.as_ref().unwrap();
        assert_eq!(labels.position, Some(DataLabelPosition::Top));
        assert!(labels.show_value);
        assert!(labels.show_series_name);
        assert!(labels.show_leader_lines);
        assert!(
            std::str::from_utf8(labels.shape_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("DDEEFF")
        );
        assert!(
            std::str::from_utf8(labels.text_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("1200000")
        );
        assert!(
            std::str::from_utf8(
                labels
                    .leader_lines
                    .as_ref()
                    .unwrap()
                    .shape_properties
                    .as_ref()
                    .unwrap()
                    .as_xml()
            )
            .unwrap()
            .contains("38100")
        );
        assert!(labels.extension_list.is_some());
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
        assert!(point_label.shape_properties.is_some());
        assert!(
            std::str::from_utf8(point_label.text_properties.as_ref().unwrap().as_xml())
                .unwrap()
                .contains("vert")
        );
        assert!(point_label.extension_list.is_some());
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
        category.min = Some(1.0);
        category.max = Some(12.0);
        category.log_base = Some(2.0);
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
        value.display_units = Some(Box::new(display_units));

        let mut date = DateAxis::new(30, AxisPosition::Top, 40);
        date.min = Some(45_000.0);
        date.max = Some(46_000.0);
        date.log_base = Some(10.0);
        date.major_unit = Some(2.0);
        date.minor_unit = Some(1.0);
        date.major_time_unit = Some(TimeUnit::Months);
        date.minor_time_unit = Some(TimeUnit::Days);
        date.base_time_unit = Some(TimeUnit::Years);
        date.auto = false;
        date.label_offset = Some(175);

        let mut series = SeriesAxis::new(40, AxisPosition::Right, 30);
        series.min = Some(1.0);
        series.max = Some(8.0);
        series.log_base = Some(2.0);
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
        assert_eq!(category.min, Some(1.0));
        assert_eq!(category.max, Some(12.0));
        assert_eq!(category.log_base, Some(2.0));
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
        assert_eq!(date.log_base, Some(10.0));
        assert_eq!(date.major_time_unit, Some(TimeUnit::Months));
        assert_eq!(date.minor_time_unit, Some(TimeUnit::Days));
        assert_eq!(date.base_time_unit, Some(TimeUnit::Years));
        assert!(!date.auto);
        assert_eq!(date.label_offset, Some(175));

        let Axis::Series(series) = &parsed.plot_area.axes[3] else {
            unreachable!();
        };
        assert_eq!(series.tick_label_skip, Some(4));
        assert_eq!(series.tick_mark_skip, Some(5));
        assert_eq!(series.min, Some(1.0));
        assert_eq!(series.max, Some(8.0));
        assert_eq!(series.log_base, Some(2.0));
    }

    #[test]
    fn rejects_invalid_scaling_on_every_axis_kind() {
        for axis in [
            r#"<c:catAx><c:axId val="1"/><c:scaling><c:min val="1"/><c:min val="2"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>"#,
            r#"<c:valAx><c:axId val="1"/><c:scaling><c:max val="1"/><c:min val="2"/></c:scaling><c:axPos val="l"/><c:crossAx val="2"/></c:valAx>"#,
            r#"<c:dateAx><c:axId val="1"/><c:scaling><c:logBase val="1"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:dateAx>"#,
            r#"<c:serAx><c:axId val="1"/><c:scaling><c:logBase val="1001"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:serAx>"#,
        ] {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>{axis}</c:plotArea></c:chart></c:chartSpace>"#
            );
            assert!(parse_chart(xml.as_bytes()).is_err());
        }
    }
}
