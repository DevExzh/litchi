//! Lossless axis and series mutation for standalone and embedded chart documents.

use super::authoring::{serialize_chart_axis_fragment, serialize_chart_series_fragment};
use super::{Axis, AxisSpec, Dimension, Document, Series, SeriesSpec};
use litchi_core::{Error, Result, xml::escape_xml};
use litchi_odf_common::chart::read;
use quick_xml::{
    Reader,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

const CHART_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const XML_NS: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MAX_CONTENT: usize = 16 * 1_048_576;
const MAX_DEPTH: usize = 128;
const MAX_ITEMS: usize = 65_536;

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct AxisUpdate {
    pub dimension: Option<Dimension>,
    pub name: Option<Option<String>>,
    pub style_name: Option<Option<String>>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SeriesUpdate {
    pub xml_id: Option<Option<String>>,
    pub class: Option<Option<String>>,
    pub values_cell_range_address: Option<Option<String>>,
    pub label_cell_address: Option<Option<String>>,
    pub attached_axis: Option<Option<String>>,
    pub style_name: Option<Option<String>>,
}

#[derive(Clone)]
struct Span {
    start: usize,
    end: usize,
}
#[derive(Clone)]
struct AxisEntry {
    span: Span,
    dimension: Dimension,
    name: Option<String>,
    style: Option<String>,
}
#[derive(Clone)]
struct SeriesEntry {
    span: Span,
    xml_id: Option<String>,
    class: Option<String>,
    values: Option<String>,
    label: Option<String>,
    axis: Option<String>,
    style: Option<String>,
}
struct PlotScan {
    axes: Vec<AxisEntry>,
    series: Vec<SeriesEntry>,
    axis_insert: usize,
    series_insert: usize,
    empty_plot: Option<Span>,
}

impl Document {
    /// Find an axis by its plot-area-unique name.
    pub fn find_axis(&self, name: &str) -> Option<Axis<'_>> {
        self.plot_area()?
            .axes()
            .find(|axis| axis.name() == Some(name))
    }

    /// Find a series by its document-unique `xml:id`.
    pub fn find_series(&self, xml_id: &str) -> Option<Series<'_>> {
        self.plot_area()?
            .series()
            .find(|series| series.xml_id() == Some(xml_id))
    }

    pub fn add_axis(&mut self, axis: &AxisSpec) -> Result<usize> {
        let xml = self.package.content_xml()?;
        let scan = scan_plot(&xml)?;
        let updated = insert_into_plot(
            &xml,
            &scan,
            scan.axis_insert,
            &serialize_chart_axis_fragment(axis)?,
        )?;
        let index = scan.axes.len();
        self.commit_chart_xml(updated)?;
        Ok(index)
    }

    pub fn replace_axis(&mut self, index: usize, axis: &AxisSpec) -> Result<()> {
        let xml = self.package.content_xml()?;
        let scan = scan_plot(&xml)?;
        let span = &scan
            .axes
            .get(index)
            .ok_or_else(|| bounds("axis", index, scan.axes.len()))?
            .span;
        self.commit_chart_xml(splice(
            &xml,
            span.start,
            span.end,
            &serialize_chart_axis_fragment(axis)?,
        )?)
    }

    pub fn update_axis(&mut self, index: usize, update: &AxisUpdate) -> Result<()> {
        let xml = self.package.content_xml()?;
        let scan = scan_plot(&xml)?;
        let entry = scan
            .axes
            .get(index)
            .ok_or_else(|| bounds("axis", index, scan.axes.len()))?;
        let mut axis = AxisSpec::new(update.dimension.unwrap_or(entry.dimension));
        axis.name = update.name.clone().unwrap_or_else(|| entry.name.clone());
        axis.style_name = update
            .style_name
            .clone()
            .unwrap_or_else(|| entry.style.clone());
        let replacement = serialize_chart_axis_fragment(&axis)?;
        let original = start_tag(&xml[entry.span.start..entry.span.end])?;
        let end = entry.span.start + original.len();
        let merged = merge_start(
            original,
            start_tag(&replacement)?,
            &["dimension", "name", "style-name"],
        )?;
        self.commit_chart_xml(splice(&xml, entry.span.start, end, &merged)?)
    }

    pub fn remove_axis(&mut self, index: usize) -> Result<()> {
        let xml = self.package.content_xml()?;
        let scan = scan_plot(&xml)?;
        let span = &scan
            .axes
            .get(index)
            .ok_or_else(|| bounds("axis", index, scan.axes.len()))?
            .span;
        self.commit_chart_xml(splice(&xml, span.start, span.end, "")?)
    }

    pub fn reorder_axis(&mut self, from: usize, to: usize) -> Result<()> {
        let xml = self.package.content_xml()?;
        let scan = scan_plot(&xml)?;
        let left = &scan
            .axes
            .get(from)
            .ok_or_else(|| bounds("axis", from, scan.axes.len()))?
            .span;
        let right = &scan
            .axes
            .get(to)
            .ok_or_else(|| bounds("axis", to, scan.axes.len()))?
            .span;
        self.commit_chart_xml(swap(&xml, left, right)?)
    }

    pub fn add_series(&mut self, series: &SeriesSpec) -> Result<usize> {
        let xml = self.package.content_xml()?;
        let scan = scan_plot(&xml)?;
        let updated = insert_into_plot(
            &xml,
            &scan,
            scan.series_insert,
            &serialize_chart_series_fragment(series)?,
        )?;
        let index = scan.series.len();
        self.commit_chart_xml(updated)?;
        Ok(index)
    }

    pub fn replace_series(&mut self, index: usize, series: &SeriesSpec) -> Result<()> {
        let xml = self.package.content_xml()?;
        let scan = scan_plot(&xml)?;
        let span = &scan
            .series
            .get(index)
            .ok_or_else(|| bounds("series", index, scan.series.len()))?
            .span;
        self.commit_chart_xml(splice(
            &xml,
            span.start,
            span.end,
            &serialize_chart_series_fragment(series)?,
        )?)
    }

    pub fn update_series(&mut self, index: usize, update: &SeriesUpdate) -> Result<()> {
        let xml = self.package.content_xml()?;
        let scan = scan_plot(&xml)?;
        let entry = scan
            .series
            .get(index)
            .ok_or_else(|| bounds("series", index, scan.series.len()))?;
        let series = SeriesSpec {
            xml_id: update
                .xml_id
                .clone()
                .unwrap_or_else(|| entry.xml_id.clone()),
            class: update.class.clone().unwrap_or_else(|| entry.class.clone()),
            values_cell_range_address: update
                .values_cell_range_address
                .clone()
                .unwrap_or_else(|| entry.values.clone()),
            label_cell_address: update
                .label_cell_address
                .clone()
                .unwrap_or_else(|| entry.label.clone()),
            attached_axis: update
                .attached_axis
                .clone()
                .unwrap_or_else(|| entry.axis.clone()),
            style_name: update
                .style_name
                .clone()
                .unwrap_or_else(|| entry.style.clone()),
            ..SeriesSpec::default()
        };
        let replacement = serialize_chart_series_fragment(&series)?;
        let original = start_tag(&xml[entry.span.start..entry.span.end])?;
        let end = entry.span.start + original.len();
        let merged = merge_start(
            original,
            start_tag(&replacement)?,
            &[
                "id",
                "class",
                "values-cell-range-address",
                "label-cell-address",
                "attached-axis",
                "style-name",
            ],
        )?;
        self.commit_chart_xml(splice(&xml, entry.span.start, end, &merged)?)
    }

    pub fn remove_series(&mut self, index: usize) -> Result<()> {
        let xml = self.package.content_xml()?;
        let scan = scan_plot(&xml)?;
        let span = &scan
            .series
            .get(index)
            .ok_or_else(|| bounds("series", index, scan.series.len()))?
            .span;
        self.commit_chart_xml(splice(&xml, span.start, span.end, "")?)
    }

    pub fn reorder_series(&mut self, from: usize, to: usize) -> Result<()> {
        let xml = self.package.content_xml()?;
        let scan = scan_plot(&xml)?;
        let left = &scan
            .series
            .get(from)
            .ok_or_else(|| bounds("series", from, scan.series.len()))?
            .span;
        let right = &scan
            .series
            .get(to)
            .ok_or_else(|| bounds("series", to, scan.series.len()))?
            .span;
        self.commit_chart_xml(swap(&xml, left, right)?)
    }

    fn commit_chart_xml(&mut self, xml: String) -> Result<()> {
        let parsed = read(&xml)?;
        scan_plot(&xml)?;
        self.package.replace_content_xml(xml)?;
        self.chart = parsed;
        Ok(())
    }
}

fn scan_plot(xml: &str) -> Result<PlotScan> {
    if xml.len() > MAX_CONTENT {
        return invalid("chart content exceeds 16 MiB mutation limit");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut plot_depth = None;
    let mut plot_close = None;
    let mut empty_plot = None;
    let mut active: Option<(bool, usize, usize, AxisOrSeries)> = None;
    let mut axes = Vec::new();
    let mut series = Vec::new();
    let mut axis_tail = None;
    let mut series_tail = None;
    loop {
        let start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let chart_ns = matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == CHART_NS);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                let local = element.local_name();
                if chart_ns && local.as_ref() == b"plot-area" {
                    if plot_depth.is_some() || plot_close.is_some() {
                        return invalid("duplicate or nested chart:plot-area");
                    }
                    plot_depth = Some(depth);
                } else if plot_depth.is_some_and(|value| depth == value + 1)
                    && chart_ns
                    && local.as_ref() == b"axis"
                {
                    active = Some((
                        true,
                        depth,
                        start,
                        AxisOrSeries::Axis(axis_from_start(&reader, &element)?),
                    ));
                } else if plot_depth.is_some_and(|value| depth == value + 1)
                    && chart_ns
                    && local.as_ref() == b"series"
                {
                    axis_tail.get_or_insert(start);
                    active = Some((
                        false,
                        depth,
                        start,
                        AxisOrSeries::Series(series_from_start(&reader, &element)?),
                    ));
                } else if plot_depth.is_some_and(|value| depth == value + 1)
                    && chart_ns
                    && is_series_tail(local.as_ref())
                {
                    axis_tail.get_or_insert(start);
                    series_tail.get_or_insert(start);
                } else if plot_depth.is_some_and(|value| depth == value + 1) {
                    // Keep schema-defined and extension tails after newly inserted axes/series.
                    axis_tail.get_or_insert(start);
                    series_tail.get_or_insert(start);
                }
                depth += 1;
                if depth > MAX_DEPTH {
                    return invalid("chart mutation XML exceeds depth limit");
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if chart_ns && local.as_ref() == b"plot-area" {
                    if plot_depth.is_some() || plot_close.is_some() {
                        return invalid("duplicate or nested chart:plot-area");
                    }
                    empty_plot = Some(Span { start, end });
                    plot_close = Some(start);
                } else if plot_depth.is_some_and(|value| depth == value + 1)
                    && chart_ns
                    && local.as_ref() == b"axis"
                {
                    let mut value = axis_from_start(&reader, &element)?;
                    value.span = Span { start, end };
                    axes.push(value);
                } else if plot_depth.is_some_and(|value| depth == value + 1)
                    && chart_ns
                    && local.as_ref() == b"series"
                {
                    axis_tail.get_or_insert(start);
                    let mut value = series_from_start(&reader, &element)?;
                    value.span = Span { start, end };
                    series.push(value);
                } else if plot_depth.is_some_and(|value| depth == value + 1) {
                    axis_tail.get_or_insert(start);
                    series_tail.get_or_insert(start);
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| make_error("chart XML depth underflow"))?;
                if active
                    .as_ref()
                    .is_some_and(|(_, value, _, _)| *value == depth)
                {
                    let (_, _, item_start, value) = active.take().expect("active chart item");
                    match value {
                        AxisOrSeries::Axis(mut value) => {
                            value.span = Span {
                                start: item_start,
                                end,
                            };
                            axes.push(value);
                        },
                        AxisOrSeries::Series(mut value) => {
                            value.span = Span {
                                start: item_start,
                                end,
                            };
                            series.push(value);
                        },
                    }
                } else if plot_depth == Some(depth)
                    && chart_ns
                    && element.local_name().as_ref() == b"plot-area"
                {
                    plot_close = Some(start);
                    plot_depth = None;
                }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in chart mutation XML"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if axes.len() > MAX_ITEMS || series.len() > MAX_ITEMS {
        return invalid("chart axis or series count exceeds limit");
    }
    let close = plot_close.ok_or_else(|| make_error("chart has no closed plot area"))?;
    axes.sort_by_key(|value| value.span.start);
    series.sort_by_key(|value| value.span.start);
    validate_references(&axes, &series)?;
    Ok(PlotScan {
        axes,
        series,
        axis_insert: axis_tail.unwrap_or(close),
        series_insert: series_tail.unwrap_or(close),
        empty_plot,
    })
}

enum AxisOrSeries {
    Axis(AxisEntry),
    Series(SeriesEntry),
}
fn axis_from_start(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<AxisEntry> {
    let dimension = match chart_attr(reader, element, b"dimension")?.as_deref() {
        Some("x") => Dimension::X,
        Some("y") => Dimension::Y,
        Some("z") => Dimension::Z,
        Some(value) => return invalid(format!("invalid chart axis dimension '{value}'")),
        None => return invalid("chart axis lacks dimension"),
    };
    Ok(AxisEntry {
        span: Span { start: 0, end: 0 },
        dimension,
        name: chart_attr(reader, element, b"name")?,
        style: chart_attr(reader, element, b"style-name")?,
    })
}
fn series_from_start(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<SeriesEntry> {
    Ok(SeriesEntry {
        span: Span { start: 0, end: 0 },
        xml_id: ns_attr(reader, element, XML_NS, b"id")?,
        class: chart_attr(reader, element, b"class")?,
        values: chart_attr(reader, element, b"values-cell-range-address")?,
        label: chart_attr(reader, element, b"label-cell-address")?,
        axis: chart_attr(reader, element, b"attached-axis")?,
        style: chart_attr(reader, element, b"style-name")?,
    })
}
fn validate_references(axes: &[AxisEntry], series: &[SeriesEntry]) -> Result<()> {
    let mut names = HashSet::new();
    for axis in axes {
        if let Some(name) = &axis.name
            && !names.insert(name.as_str())
        {
            return invalid(format!("duplicate chart axis name '{name}'"));
        }
    }
    let mut ids = HashSet::new();
    for value in series {
        if let Some(id) = &value.xml_id
            && !ids.insert(id.as_str())
        {
            return invalid(format!("duplicate chart series xml:id '{id}'"));
        }
        if let Some(axis) = &value.axis
            && !names.contains(axis.as_str())
        {
            return invalid(format!("series references unknown axis '{axis}'"));
        }
    }
    Ok(())
}
fn is_series_tail(local: &[u8]) -> bool {
    matches!(
        local,
        b"stock-gain-marker" | b"stock-loss-marker" | b"stock-range-line" | b"wall" | b"floor"
    )
}
fn chart_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<Option<String>> {
    ns_attr(reader, element, CHART_NS, local)
}
fn ns_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if matches!(reader.resolver().resolve_attribute(attribute.key), (ResolveResult::Bound(Namespace(uri)), name) if uri == namespace && name.as_ref() == local)
        {
            return attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(xml_error);
        }
    }
    Ok(None)
}
fn start_tag(xml: &str) -> Result<&str> {
    xml.find('>')
        .map(|value| &xml[..=value])
        .ok_or_else(|| make_error("unterminated chart start tag"))
}
fn merge_start(original: &str, replacement: &str, known: &[&str]) -> Result<String> {
    let (name, old) = parse_start(original)?;
    let (_, mut attrs) = parse_start(replacement)?;
    let prefix = name.split_once(':').map_or("", |value| value.0);
    for (attr, value) in old {
        let (attr_prefix, local) = attr.split_once(':').unwrap_or(("", attr.as_str()));
        let standard = (attr_prefix == prefix || attr_prefix == "chart" || attr_prefix == "xml")
            && known.contains(&local);
        if !standard && !attrs.iter().any(|value| value.0 == attr) {
            attrs.push((attr, value));
        }
    }
    let mut out = format!("<{name}");
    for (name, value) in attrs {
        out.push(' ');
        out.push_str(&name);
        out.push_str("=\"");
        out.push_str(&escape_xml(&value));
        out.push('"');
    }
    if original.trim_end().ends_with("/>") {
        out.push_str("/>");
    } else {
        out.push('>');
    }
    Ok(out)
}
fn parse_start(xml: &str) -> Result<(String, Vec<(String, String)>)> {
    let mut reader = Reader::from_str(xml);
    let decoder = reader.decoder();
    match reader.read_event().map_err(xml_error)? {
        Event::Start(value) | Event::Empty(value) => {
            let name = std::str::from_utf8(value.name().as_ref())
                .map_err(|_| make_error("invalid chart QName"))?
                .to_string();
            let mut attrs = Vec::new();
            for attr in value.attributes().with_checks(true) {
                let attr = attr.map_err(xml_error)?;
                attrs.push((
                    std::str::from_utf8(attr.key.as_ref())
                        .map_err(|_| make_error("invalid chart attribute QName"))?
                        .to_string(),
                    attr.decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                        .map_err(xml_error)?
                        .into_owned(),
                ));
            }
            Ok((name, attrs))
        },
        _ => invalid("expected chart start tag"),
    }
}
fn insert_into_plot(xml: &str, scan: &PlotScan, offset: usize, fragment: &str) -> Result<String> {
    if let Some(span) = &scan.empty_plot {
        let original = &xml[span.start..span.end];
        let trimmed = original.trim_end();
        let head = trimmed
            .strip_suffix("/>")
            .ok_or_else(|| make_error("invalid empty chart:plot-area"))?;
        let (name, _) = parse_start(trimmed)?;
        return splice(
            xml,
            span.start,
            span.end,
            &format!("{head}>{fragment}</{name}>"),
        );
    }
    splice(xml, offset, offset, fragment)
}
fn swap(xml: &str, a: &Span, b: &Span) -> Result<String> {
    if a.start == b.start {
        return Ok(xml.to_string());
    }
    let (a, b) = if a.start < b.start { (a, b) } else { (b, a) };
    if a.end > b.start {
        return invalid("overlapping chart spans");
    }
    Ok(format!(
        "{}{}{}{}{}",
        &xml[..a.start],
        &xml[b.start..b.end],
        &xml[a.end..b.start],
        &xml[a.start..a.end],
        &xml[b.end..]
    ))
}
fn splice(xml: &str, start: usize, end: usize, value: &str) -> Result<String> {
    if start > end || end > xml.len() || !xml.is_char_boundary(start) || !xml.is_char_boundary(end)
    {
        return invalid("invalid chart mutation span");
    }
    Ok(format!("{}{}{}", &xml[..start], value, &xml[end..]))
}
fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_| make_error("chart XML position overflow"))
}
fn bounds(kind: &str, index: usize, len: usize) -> Error {
    make_error(format!(
        "chart {kind} index {index} is out of bounds for {len} entries"
    ))
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    make_error(format!("invalid chart mutation XML: {error}"))
}
fn make_error(value: impl Into<String>) -> Error {
    Error::InvalidFormat(value.into())
}
fn invalid<T>(value: impl Into<String>) -> Result<T> {
    Err(make_error(value))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Definition, StyleElement};

    #[test]
    fn mixed_and_three_dimensional_chart_mutations_are_atomic() {
        let mut definition = Definition::new("chart:bar");
        let mut x = AxisSpec::new(Dimension::X);
        x.name = Some("x".to_string());
        let mut y = AxisSpec::new(Dimension::Y);
        y.name = Some("y".to_string());
        definition.plot_area.axes = vec![x, y];
        definition.plot_area.wall = Some(StyleElement::default());
        definition.plot_area.floor = Some(StyleElement::default());
        definition.plot_area.stock_range_line = Some(StyleElement::default());
        definition.plot_area.series.push(SeriesSpec {
            xml_id: Some("bars".to_string()),
            class: Some("chart:bar".to_string()),
            attached_axis: Some("y".to_string()),
            ..Default::default()
        });
        let mut document = Document::create(&definition).unwrap();
        document
            .add_series(&SeriesSpec {
                xml_id: Some("line".to_string()),
                class: Some("chart:line".to_string()),
                attached_axis: Some("y".to_string()),
                ..Default::default()
            })
            .unwrap();
        document.reorder_series(0, 1).unwrap();
        assert!(document.find_series("line").is_some());
        assert!(document.remove_axis(1).is_err());
        document
            .update_series(
                0,
                &SeriesUpdate {
                    style_name: Some(Some("seriesStyle".to_string())),
                    ..Default::default()
                },
            )
            .unwrap();
    }
}
