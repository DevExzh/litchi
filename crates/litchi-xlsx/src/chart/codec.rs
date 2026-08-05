//! DrawingML chart and user-shapes XML boundaries.

use super::anchor::Anchor;
use super::model::Chart;
use crate::error::{Error, Result};
use litchi_drawingml::chart::{axis::Axis, model::Chart as ChartModel, plot_area::TypeGroup};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

/// Parse a chart from chart XML and its worksheet anchor.
pub fn parse_chart_from_xml(chart_xml: &[u8], anchor: Anchor) -> Result<Chart> {
    let chart = litchi_drawingml::chart::reader::read(chart_xml)?;
    Ok(Chart::new(chart, anchor))
}

/// Generate chart XML for a worksheet chart.
pub fn generate_chart_xml(chart: &ChartModel) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    litchi_drawingml::chart::writer::write(&mut output, chart)
        .map_err(|e| Error::Invalid(e.to_string()))?;
    Ok(output)
}

pub(crate) fn generate_chart_xml_with_external_data_id(
    chart: &ChartModel,
    external_data_relationship_id: Option<&str>,
    user_shapes_relationship_id: Option<&str>,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    litchi_drawingml::chart::writer::write_with_rels(
        &mut output,
        chart,
        external_data_relationship_id,
        user_shapes_relationship_id,
    )
    .map_err(|e| Error::Invalid(e.to_string()))?;
    Ok(output)
}

pub(crate) fn chart_user_shapes_relationship_ids(xml: &[u8]) -> Result<HashSet<String>> {
    const RELATIONSHIPS_NAMESPACE: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
        b"http://purl.oclc.org/ooxml/officeDocument/relationships";

    let xml = litchi_ooxml_common::mce::process_ooxml(xml)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut closed_root = false;
    let mut relationship_ids = HashSet::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Invalid(error.to_string()))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                if depth == 0 {
                    if saw_root
                        || !litchi_ooxml_common::xml::is_drawingml_chart_name(
                            &namespace,
                            element.name(),
                            b"userShapes",
                        )
                    {
                        return Err(Error::Invalid(
                            "chart user-shapes XML must have one chart userShapes root".into(),
                        ));
                    }
                    saw_root = true;
                }
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Invalid(error.to_string()))?;
                    let (attribute_namespace, _) =
                        reader.resolver().resolve_attribute(attribute.key);
                    if matches!(
                        attribute_namespace,
                        ResolveResult::Bound(Namespace(value))
                            if value == RELATIONSHIPS_NAMESPACE
                                || value == STRICT_RELATIONSHIPS_NAMESPACE
                    ) {
                        relationship_ids.insert(
                            attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Explicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| Error::Invalid(error.to_string()))?
                                .into_owned(),
                        );
                    }
                }
                if matches!(event, Event::Start(_)) {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::Invalid("chart user-shapes XML nesting is too deep".into())
                    })?;
                } else if depth == 0 {
                    closed_root = true;
                }
            },
            Event::End(ref element) => {
                if depth == 0 {
                    return Err(Error::Invalid(
                        "chart user-shapes XML has an unmatched closing element".into(),
                    ));
                }
                depth -= 1;
                if depth == 0 {
                    if !litchi_ooxml_common::xml::is_drawingml_chart_name(
                        &namespace,
                        element.name(),
                        b"userShapes",
                    ) {
                        return Err(Error::Invalid(
                            "chart user-shapes XML has an invalid root closing element".into(),
                        ));
                    }
                    closed_root = true;
                }
            },
            Event::Text(ref text)
                if depth == 0
                    && !text
                        .decode()
                        .map_err(|error| Error::Invalid(error.to_string()))?
                        .trim()
                        .is_empty() =>
            {
                return Err(Error::Invalid(
                    "chart user-shapes XML contains text outside its root".into(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::Invalid(
                    "chart user-shapes XML contains data outside its root".into(),
                ));
            },
            Event::DocType(_) => {
                return Err(Error::Invalid(
                    "chart user-shapes XML cannot contain a document type".into(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !saw_root || !closed_root || depth != 0 {
        return Err(Error::Invalid(
            "chart user-shapes XML has no complete root".into(),
        ));
    }
    Ok(relationship_ids)
}

fn append_chart_line_fragment<'a>(
    fragments: &mut Vec<&'a [u8]>,
    lines: Option<&'a litchi_drawingml::chart::Lines>,
) {
    if let Some(shape_properties) = lines.and_then(|lines| lines.shape_properties.as_ref()) {
        fragments.push(shape_properties.as_xml());
    }
}

fn append_up_down_bar_fragments<'a>(
    fragments: &mut Vec<&'a [u8]>,
    bars: Option<&'a litchi_drawingml::chart::UpDownBars>,
) {
    let Some(bars) = bars else {
        return;
    };
    append_chart_line_fragment(fragments, bars.up_bars.as_ref());
    append_chart_line_fragment(fragments, bars.down_bars.as_ref());
    if let Some(extension_list) = bars.extension_list.as_ref() {
        fragments.push(extension_list.as_xml());
    }
}

fn append_point_data_label_fragments<'a>(
    fragments: &mut Vec<&'a [u8]>,
    label: Option<&'a litchi_drawingml::chart::DataLabel>,
) {
    let Some(label) = label else {
        return;
    };
    if let Some(shape_properties) = label.shape_properties.as_ref() {
        fragments.push(shape_properties.as_xml());
    }
    if let Some(text_properties) = label.text_properties.as_ref() {
        fragments.push(text_properties.as_xml());
    }
    if let Some(extension_list) = label.extension_list.as_ref() {
        fragments.push(extension_list.as_xml());
    }
}

fn append_data_label_fragments<'a>(
    fragments: &mut Vec<&'a [u8]>,
    labels: Option<&'a litchi_drawingml::chart::DataLabels>,
) {
    let Some(labels) = labels else {
        return;
    };
    if let Some(shape_properties) = labels.shape_properties.as_ref() {
        fragments.push(shape_properties.as_xml());
    }
    if let Some(text_properties) = labels.text_properties.as_ref() {
        fragments.push(text_properties.as_xml());
    }
    append_chart_line_fragment(fragments, labels.leader_lines.as_ref());
    if let Some(extension_list) = labels.extension_list.as_ref() {
        fragments.push(extension_list.as_xml());
    }
    for label in &labels.labels {
        append_point_data_label_fragments(fragments, Some(label));
    }
}

pub(crate) fn chart_fragment_relationship_ids(chart: &ChartModel) -> Result<HashSet<String>> {
    const RELATIONSHIPS_NAMESPACE: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
        b"http://purl.oclc.org/ooxml/officeDocument/relationships";

    let mut relationship_ids = HashSet::new();
    let mut fragments = Vec::new();
    fragments.extend(
        [
            chart
                .shape_properties
                .as_ref()
                .map(litchi_drawingml::chart::ShapeProperties::as_xml),
            chart
                .text_properties
                .as_ref()
                .map(litchi_drawingml::chart::TextProperties::as_xml),
            chart
                .extension_list
                .as_ref()
                .map(litchi_drawingml::chart::ExtensionList::as_xml),
            chart
                .chart_extension_list
                .as_ref()
                .map(litchi_drawingml::chart::ExtensionList::as_xml),
            chart
                .title
                .as_ref()
                .and(chart.title_shape_properties.as_ref())
                .map(litchi_drawingml::chart::ShapeProperties::as_xml),
            chart
                .title
                .as_ref()
                .and(chart.title_text_properties.as_ref())
                .map(litchi_drawingml::chart::TextProperties::as_xml),
            chart
                .title
                .as_ref()
                .and(chart.title_extension_list.as_ref())
                .map(litchi_drawingml::chart::ExtensionList::as_xml),
            chart
                .plot_area
                .shape_properties
                .as_ref()
                .map(litchi_drawingml::chart::ShapeProperties::as_xml),
            chart
                .plot_area
                .extension_list
                .as_ref()
                .map(litchi_drawingml::chart::ExtensionList::as_xml),
            chart
                .plot_area
                .data_table
                .as_ref()
                .and_then(|table| table.shape_properties.as_ref())
                .map(litchi_drawingml::chart::ShapeProperties::as_xml),
            chart
                .plot_area
                .data_table
                .as_ref()
                .and_then(|table| table.text_properties.as_ref())
                .map(litchi_drawingml::chart::TextProperties::as_xml),
            chart
                .plot_area
                .data_table
                .as_ref()
                .and_then(|table| table.extension_list.as_ref())
                .map(litchi_drawingml::chart::ExtensionList::as_xml),
            chart
                .floor
                .as_ref()
                .and_then(|surface| surface.shape_properties.as_ref())
                .map(litchi_drawingml::chart::ShapeProperties::as_xml),
            chart
                .back_wall
                .as_ref()
                .and_then(|surface| surface.shape_properties.as_ref())
                .map(litchi_drawingml::chart::ShapeProperties::as_xml),
            chart
                .side_wall
                .as_ref()
                .and_then(|surface| surface.shape_properties.as_ref())
                .map(litchi_drawingml::chart::ShapeProperties::as_xml),
            chart
                .floor
                .as_ref()
                .and_then(|surface| surface.extension_list.as_ref())
                .map(litchi_drawingml::chart::ExtensionList::as_xml),
            chart
                .back_wall
                .as_ref()
                .and_then(|surface| surface.extension_list.as_ref())
                .map(litchi_drawingml::chart::ExtensionList::as_xml),
            chart
                .side_wall
                .as_ref()
                .and_then(|surface| surface.extension_list.as_ref())
                .map(litchi_drawingml::chart::ExtensionList::as_xml),
            chart
                .legend
                .as_ref()
                .and_then(|legend| legend.shape_properties.as_ref())
                .map(litchi_drawingml::chart::ShapeProperties::as_xml),
            chart
                .legend
                .as_ref()
                .and_then(|legend| legend.text_properties.as_ref())
                .map(litchi_drawingml::chart::TextProperties::as_xml),
            chart
                .legend
                .as_ref()
                .and_then(|legend| legend.extension_list.as_ref())
                .map(litchi_drawingml::chart::ExtensionList::as_xml),
        ]
        .into_iter()
        .flatten(),
    );
    if let Some(legend) = chart.legend.as_ref() {
        for entry in &legend.entries {
            fragments.extend(
                [
                    entry
                        .text_properties
                        .as_ref()
                        .map(litchi_drawingml::chart::TextProperties::as_xml),
                    entry
                        .extension_list
                        .as_ref()
                        .map(litchi_drawingml::chart::ExtensionList::as_xml),
                ]
                .into_iter()
                .flatten(),
            );
        }
    }
    if let Some(formats) = chart.pivot_formats.as_ref() {
        for format in formats {
            fragments.extend(
                [
                    format
                        .shape_properties
                        .as_ref()
                        .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                    format
                        .text_properties
                        .as_ref()
                        .map(litchi_drawingml::chart::TextProperties::as_xml),
                    format
                        .extension_list
                        .as_ref()
                        .map(litchi_drawingml::chart::ExtensionList::as_xml),
                    format
                        .marker
                        .as_ref()
                        .and_then(|marker| marker.shape_properties.as_ref())
                        .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                    format
                        .marker
                        .as_ref()
                        .and_then(|marker| marker.extension_list.as_ref())
                        .map(litchi_drawingml::chart::ExtensionList::as_xml),
                ]
                .into_iter()
                .flatten(),
            );
            append_point_data_label_fragments(&mut fragments, format.data_label.as_ref());
        }
    }
    for axis in &chart.plot_area.axes {
        let common = axis.common();
        fragments.extend(
            [
                common
                    .title
                    .as_ref()
                    .and(common.title_shape_properties.as_ref())
                    .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                common
                    .title
                    .as_ref()
                    .and(common.title_text_properties.as_ref())
                    .map(litchi_drawingml::chart::TextProperties::as_xml),
                common
                    .title
                    .as_ref()
                    .and(common.title_extension_list.as_ref())
                    .map(litchi_drawingml::chart::ExtensionList::as_xml),
                common
                    .major_gridlines
                    .as_ref()
                    .and_then(|lines| lines.shape_properties.as_ref())
                    .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                common
                    .minor_gridlines
                    .as_ref()
                    .and_then(|lines| lines.shape_properties.as_ref())
                    .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                common
                    .shape_properties
                    .as_ref()
                    .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                common
                    .text_properties
                    .as_ref()
                    .map(litchi_drawingml::chart::TextProperties::as_xml),
                common
                    .scaling_extension_list
                    .as_ref()
                    .map(litchi_drawingml::chart::ExtensionList::as_xml),
                common
                    .extension_list
                    .as_ref()
                    .map(litchi_drawingml::chart::ExtensionList::as_xml),
            ]
            .into_iter()
            .flatten(),
        );
        if let Axis::Value(axis) = axis
            && let Some(display_units) = axis.display_units.as_ref()
        {
            fragments.extend(
                [
                    display_units
                        .label_shape_properties
                        .as_ref()
                        .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                    display_units
                        .label_text_properties
                        .as_ref()
                        .map(litchi_drawingml::chart::TextProperties::as_xml),
                    display_units
                        .extension_list
                        .as_ref()
                        .map(litchi_drawingml::chart::ExtensionList::as_xml),
                ]
                .into_iter()
                .flatten(),
            );
        }
    }
    for group in &chart.plot_area.type_groups {
        match group {
            TypeGroup::Area(group) => {
                append_chart_line_fragment(&mut fragments, group.drop_lines.as_ref());
            },
            TypeGroup::Area3D(group) => {
                append_chart_line_fragment(&mut fragments, group.drop_lines.as_ref());
            },
            TypeGroup::Bar(group) => {
                for lines in &group.series_lines {
                    append_chart_line_fragment(&mut fragments, Some(lines));
                }
            },
            TypeGroup::Line(group) => {
                append_chart_line_fragment(&mut fragments, group.drop_lines.as_ref());
                append_chart_line_fragment(&mut fragments, group.high_low_lines.as_ref());
                append_up_down_bar_fragments(&mut fragments, group.up_down_bars.as_ref());
            },
            TypeGroup::Line3D(group) => {
                append_chart_line_fragment(&mut fragments, group.drop_lines.as_ref());
            },
            TypeGroup::OfPie(group) => {
                for lines in &group.series_lines {
                    append_chart_line_fragment(&mut fragments, Some(lines));
                }
            },
            TypeGroup::Stock(group) => {
                append_chart_line_fragment(&mut fragments, group.drop_lines.as_ref());
                append_chart_line_fragment(&mut fragments, group.high_low_lines.as_ref());
                append_up_down_bar_fragments(&mut fragments, group.up_down_bars.as_ref());
            },
            TypeGroup::Surface(group) => {
                if let Some(formats) = group.band_formats.as_ref() {
                    for format in formats {
                        if let Some(shape_properties) = format.shape_properties.as_ref() {
                            fragments.push(shape_properties.as_xml());
                        }
                    }
                }
            },
            TypeGroup::Surface3D(group) => {
                if let Some(formats) = group.band_formats.as_ref() {
                    for format in formats {
                        if let Some(shape_properties) = format.shape_properties.as_ref() {
                            fragments.push(shape_properties.as_xml());
                        }
                    }
                }
            },
            _ => {},
        }
        for series in &group.common().series {
            fragments.extend(
                [
                    series
                        .shape_properties
                        .as_ref()
                        .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                    series
                        .extension_list
                        .as_ref()
                        .map(litchi_drawingml::chart::ExtensionList::as_xml),
                    series
                        .marker_shape_properties
                        .as_ref()
                        .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                    series
                        .marker_extension_list
                        .as_ref()
                        .map(litchi_drawingml::chart::ExtensionList::as_xml),
                ]
                .into_iter()
                .flatten(),
            );
            append_data_label_fragments(&mut fragments, series.data_labels.as_ref());
            for error_bar in &series.error_bars {
                if let Some(shape_properties) = error_bar.shape_properties.as_ref() {
                    fragments.push(shape_properties.as_xml());
                }
                if let Some(extension_list) = error_bar.extension_list.as_ref() {
                    fragments.push(extension_list.as_xml());
                }
            }
            for trendline in &series.trendlines {
                fragments.extend(
                    [
                        trendline
                            .shape_properties
                            .as_ref()
                            .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                        trendline
                            .label_shape_properties
                            .as_ref()
                            .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                        trendline
                            .label_text_properties
                            .as_ref()
                            .map(litchi_drawingml::chart::TextProperties::as_xml),
                        trendline
                            .label_extension_list
                            .as_ref()
                            .map(litchi_drawingml::chart::ExtensionList::as_xml),
                        trendline
                            .extension_list
                            .as_ref()
                            .map(litchi_drawingml::chart::ExtensionList::as_xml),
                    ]
                    .into_iter()
                    .flatten(),
                );
            }
            for point in &series.data_points {
                fragments.extend(
                    [
                        point
                            .shape_properties
                            .as_ref()
                            .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                        point
                            .extension_list
                            .as_ref()
                            .map(litchi_drawingml::chart::ExtensionList::as_xml),
                        point
                            .marker_shape_properties
                            .as_ref()
                            .map(litchi_drawingml::chart::ShapeProperties::as_xml),
                        point
                            .marker_extension_list
                            .as_ref()
                            .map(litchi_drawingml::chart::ExtensionList::as_xml),
                    ]
                    .into_iter()
                    .flatten(),
                );
            }
        }
        if let Some(extension_list) = group.common().extension_list.as_ref() {
            fragments.push(extension_list.as_xml());
        }
    }
    for xml in fragments {
        let mut reader = NsReader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            let (_, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::Invalid(error.to_string()))?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    for attribute in element.attributes() {
                        let attribute =
                            attribute.map_err(|error| Error::Invalid(error.to_string()))?;
                        let (attribute_namespace, _) =
                            reader.resolver().resolve_attribute(attribute.key);
                        if matches!(
                            attribute_namespace,
                            ResolveResult::Bound(Namespace(value))
                                if value == RELATIONSHIPS_NAMESPACE
                                    || value == STRICT_RELATIONSHIPS_NAMESPACE
                        ) {
                            relationship_ids.insert(
                                attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Explicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| Error::Invalid(error.to_string()))?
                                    .into_owned(),
                            );
                        }
                    }
                },
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }
    }
    Ok(relationship_ids)
}
