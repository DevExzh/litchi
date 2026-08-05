//! XML discovery, validation, and shared DrawingML schema delegation.

use super::model::{
    COLOR_STYLE_CT, Conformance, Graph, MAX_ATTRIBUTE_BYTES, MAX_ATTRIBUTES, MAX_CHART_XML,
    MAX_CHARTS, MAX_COMPANION_XML, MAX_DEPTH, MAX_DOCUMENT_XML, MAX_NODES, MAX_TOTAL_BYTES,
    MAX_WORKBOOK_BYTES, R, RS, STYLE_CT,
};
use crate::error::{Error, Result};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_opc::PackURI;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::Part;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeSet, HashSet};
use std::io::Cursor;

/// Fully resolved attribute triple: `(namespace, local name, unescaped value)`.
type ResolvedAttribute = (String, String, String);
/// Fully resolved element triple: `(namespace, local name, attributes)`.
type ResolvedElement = (String, String, Vec<ResolvedAttribute>);

#[derive(Default)]
pub(crate) struct ChartScan {
    pub(crate) workbook_id: Option<String>,
    external_depth: Option<usize>,
    external_count: usize,
}

#[derive(Default)]
struct Limits {
    nodes: usize,
    attributes: usize,
    attribute_bytes: usize,
}

pub(crate) fn validate_graph_value(graph: &Graph) -> Result<()> {
    if graph.charts.len() > MAX_CHARTS {
        return Err(limit("chart count"));
    }
    let mut total = 0usize;
    let mut document_ids = HashSet::new();
    let mut parts = HashSet::new();
    for chart in &graph.charts {
        validate_id(&chart.document_relationship_id)?;
        if !document_ids.insert(chart.document_relationship_id.as_str()) {
            return Err(invalid("document chart relationship IDs collide"));
        }
        if chart.content_type != super::model::CHART_CT {
            return Err(invalid("chart has invalid content type"));
        }
        let uri = PackURI::new(&chart.part_name).map_err(Error::InvalidUri)?;
        validate_leaf_path(&uri, "/word/charts/", "chart")?;
        if !parts.insert(chart.part_name.as_str()) {
            return Err(invalid("chart resource part names collide"));
        }
        let scan = scan_chart_xml(&chart.data, graph.conformance)?;
        add_total(&mut total, chart.data.len(), MAX_CHART_XML, "chart bytes")?;
        if chart.styles.len() > super::model::MAX_COMPANIONS
            || chart.color_styles.len() > super::model::MAX_COMPANIONS
        {
            return Err(limit("chart companion count"));
        }
        let mut ids = HashSet::new();
        for (resources, content_type, root, label) in [
            (&chart.styles, STYLE_CT, "chartStyle", "chart style"),
            (
                &chart.color_styles,
                COLOR_STYLE_CT,
                "colorStyle",
                "chart color style",
            ),
        ] {
            for resource in resources {
                validate_id(&resource.relationship_id)?;
                if !ids.insert(resource.relationship_id.as_str()) {
                    return Err(invalid("chart relationship IDs collide"));
                }
                if resource.content_type != content_type {
                    return Err(invalid(format!("{label} has invalid content type")));
                }
                let uri = PackURI::new(&resource.part_name).map_err(Error::InvalidUri)?;
                validate_leaf_path(&uri, "/word/charts/", label)?;
                if !parts.insert(resource.part_name.as_str()) {
                    return Err(invalid("chart resource part names collide"));
                }
                validate_companion_xml(&resource.data, MAX_COMPANION_XML, root, label)?;
                add_total(
                    &mut total,
                    resource.data.len(),
                    MAX_COMPANION_XML,
                    "chart companion bytes",
                )?;
            }
        }
        let workbook_id = if let Some(workbook) = &chart.workbook {
            validate_id(&workbook.relationship_id)?;
            if !ids.insert(workbook.relationship_id.as_str()) {
                return Err(invalid("chart relationship IDs collide"));
            }
            let uri = PackURI::new(&workbook.part_name).map_err(Error::InvalidUri)?;
            validate_leaf_path(&uri, "/word/embeddings/", "embedded workbook")?;
            if !workbook.content_type.validates_path(uri.as_str())
                || !parts.insert(workbook.part_name.as_str())
            {
                return Err(invalid(
                    "embedded workbook path, suffix, or ownership is invalid",
                ));
            }
            add_total(
                &mut total,
                workbook.data.len(),
                MAX_WORKBOOK_BYTES,
                "embedded workbook bytes",
            )?;
            Some(workbook.relationship_id.as_str())
        } else {
            None
        };
        if scan.workbook_id.as_deref() != workbook_id {
            return Err(invalid(
                "chart externalData and embedded workbook metadata differ",
            ));
        }
    }
    Ok(())
}

pub(crate) fn document_references(xml: &[u8]) -> Result<(Conformance, Vec<String>)> {
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        if let Ok(value) = scan_document_xml(xml, conformance) {
            return Ok((conformance, value));
        }
    }
    Err(invalid("invalid DOCX document root or chart anchors"))
}

fn scan_document_xml(xml: &[u8], conformance: Conformance) -> Result<Vec<String>> {
    if xml.len() > MAX_DOCUMENT_XML {
        return Err(limit("document XML bytes"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_DOCUMENT_XML {
        return Err(limit("processed document XML bytes"));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut limits = Limits::default();
    let mut root = false;
    let mut frames: Vec<(usize, usize)> = Vec::new();
    let mut references = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                depth += 1;
                structure(&mut limits, depth)?;
                let (ns, local, attrs) = element_info(&reader, &element, &mut limits)?;
                if !root {
                    if ns != conformance.w() || local != "document" {
                        return Err(invalid("invalid document root or namespace"));
                    }
                    root = true;
                }
                if ns == conformance.a() && local == "graphicData" {
                    if attr(&attrs, "", "uri") == Some(conformance.c()) {
                        frames.push((depth, 0));
                    }
                } else if ns == conformance.c() && local == "chart" {
                    let Some(frame) = frames.last_mut() else {
                        return Err(invalid("chart element is outside chart graphicData"));
                    };
                    frame.1 += 1;
                    let id = required_rel_id(&attrs, conformance)?;
                    references.push(id.to_owned());
                }
            },
            Event::Empty(element) => {
                structure(&mut limits, depth + 1)?;
                let (ns, local, attrs) = element_info(&reader, &element, &mut limits)?;
                if !root {
                    if ns != conformance.w() || local != "document" {
                        return Err(invalid("invalid document root or namespace"));
                    }
                    root = true;
                }
                if ns == conformance.a()
                    && local == "graphicData"
                    && attr(&attrs, "", "uri") == Some(conformance.c())
                {
                    return Err(invalid("chart graphicData lacks chart child"));
                }
                if ns == conformance.c() && local == "chart" {
                    let Some(frame) = frames.last_mut() else {
                        return Err(invalid("chart element is outside chart graphicData"));
                    };
                    frame.1 += 1;
                    references.push(required_rel_id(&attrs, conformance)?.to_owned());
                }
            },
            Event::End(_) => {
                if let Some((_, count)) = frames.pop_if(|frame| frame.0 == depth)
                    && count != 1
                {
                    return Err(invalid(
                        "chart graphicData must contain exactly one chart reference",
                    ));
                }
                if depth == 0 {
                    return Err(invalid("unexpected document XML closing element"));
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected")),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root || depth != 0 || !frames.is_empty() {
        return Err(invalid("missing or unterminated document root"));
    }
    Ok(references)
}

pub(crate) fn scan_chart_xml(xml: &[u8], conformance: Conformance) -> Result<ChartScan> {
    if xml.len() > MAX_CHART_XML {
        return Err(limit("chart XML bytes"));
    }
    litchi_drawingml::chart::reader::read(Cursor::new(xml))
        .map_err(|error| invalid(format!("DrawingML chart schema: {error}")))?;
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_CHART_XML {
        return Err(limit("processed chart XML bytes"));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut limits = Limits::default();
    let mut root = false;
    let mut scan = ChartScan::default();
    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                depth += 1;
                structure(&mut limits, depth)?;
                let (ns, local, attrs) = element_info(&reader, &element, &mut limits)?;
                if !root {
                    if ns != conformance.c() || local != "chartSpace" {
                        return Err(invalid("invalid chartSpace root or namespace"));
                    }
                    if attr(&attrs, "", "fallbackImg").is_some() {
                        return Err(invalid("chart fallback image relationship is unsupported"));
                    }
                    root = true;
                }
                inspect_chart_element(&ns, &local, &attrs, conformance, depth, &mut scan)?;
            },
            Event::Empty(element) => {
                structure(&mut limits, depth + 1)?;
                let (ns, local, attrs) = element_info(&reader, &element, &mut limits)?;
                if !root {
                    if ns != conformance.c() || local != "chartSpace" {
                        return Err(invalid("invalid chartSpace root or namespace"));
                    }
                    if attr(&attrs, "", "fallbackImg").is_some() {
                        return Err(invalid("chart fallback image relationship is unsupported"));
                    }
                    root = true;
                }
                inspect_chart_element(&ns, &local, &attrs, conformance, depth + 1, &mut scan)?;
                if ns == conformance.c() && local == "externalData" {
                    scan.external_depth = None;
                }
            },
            Event::End(_) => {
                if scan.external_depth == Some(depth) {
                    scan.external_depth = None;
                }
                if depth == 0 {
                    return Err(invalid("unexpected chart XML closing element"));
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected")),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root || depth != 0 {
        return Err(invalid("missing or unterminated chartSpace root"));
    }
    Ok(scan)
}

fn inspect_chart_element(
    ns: &str,
    local: &str,
    attrs: &[ResolvedAttribute],
    conformance: Conformance,
    depth: usize,
    scan: &mut ChartScan,
) -> Result<()> {
    if ns == conformance.c() && local == "externalData" {
        scan.external_count += 1;
        if scan.external_count > 1 {
            return Err(invalid("chart has multiple externalData elements"));
        }
        scan.external_depth = Some(depth);
        scan.workbook_id = Some(required_rel_id(attrs, conformance)?.to_owned());
    } else if ns == conformance.c() && local == "autoUpdate" {
        if scan.external_depth.is_none() {
            return Err(invalid("autoUpdate is outside externalData"));
        }
        match attr(attrs, "", "val") {
            Some("0" | "false") => {},
            _ => return Err(invalid("automatic chart data updates are rejected")),
        }
    }
    for (namespace, name, _) in attrs {
        if matches!(namespace.as_str(), R | RS)
            && !(ns == conformance.c()
                && local == "externalData"
                && namespace == conformance.r()
                && name == "id")
        {
            return Err(invalid(
                "chart XML contains an unsupported relationship reference",
            ));
        }
    }
    Ok(())
}

pub(crate) fn validate_companion_xml(
    xml: &[u8],
    max: usize,
    root_name: &str,
    label: &str,
) -> Result<()> {
    if xml.len() > max {
        return Err(limit("companion XML bytes"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > max {
        return Err(limit("processed companion XML bytes"));
    }
    let result = match root_name {
        "chartStyle" => litchi_drawingml::chart::style::parse(processed.as_ref()).map(|_| ()),
        "colorStyle" => litchi_drawingml::chart::style::parse_color(processed.as_ref()).map(|_| ()),
        _ => return Err(invalid(format!("unsupported {label} root"))),
    };
    result.map_err(|error| Error::Invalid(format!("{label}: {error}")))
}

fn element_info(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: &mut Limits,
) -> Result<ResolvedElement> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let local = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    let mut values = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = item.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        limits.attributes += 1;
        if limits.attributes > MAX_ATTRIBUTES {
            return Err(limit("XML attribute count"));
        }
        let (namespace, name) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(name.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let raw_value = std::str::from_utf8(item.value.as_ref()).map_err(xml_error)?;
        let value = quick_xml::escape::unescape(raw_value)
            .map_err(xml_error)?
            .into_owned();
        limits.attribute_bytes = limits
            .attribute_bytes
            .checked_add(namespace.len() + name.len() + value.len())
            .ok_or_else(|| limit("XML attribute bytes"))?;
        if limits.attribute_bytes > MAX_ATTRIBUTE_BYTES {
            return Err(limit("XML attribute bytes"));
        }
        if values
            .iter()
            .any(|(ns, n, _): &ResolvedAttribute| ns == &namespace && n == &name)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        values.push((namespace, name, value));
    }
    Ok((namespace, local, values))
}

fn required_rel_id(attrs: &[ResolvedAttribute], conformance: Conformance) -> Result<&str> {
    let value = attr(attrs, conformance.r(), "id")
        .ok_or_else(|| invalid("chart reference lacks relationship ID"))?;
    if attrs
        .iter()
        .any(|(namespace, name, _)| !(namespace == conformance.r() && name == "id"))
    {
        return Err(invalid("chart reference has unsupported attributes"));
    }
    validate_id(value)?;
    Ok(value)
}

fn attr<'a>(attrs: &'a [ResolvedAttribute], namespace: &str, name: &str) -> Option<&'a str> {
    attrs
        .iter()
        .find(|(ns, n, _)| ns == namespace && n == name)
        .map(|(_, _, value)| value.as_str())
}

fn structure(limits: &mut Limits, depth: usize) -> Result<()> {
    limits.nodes += 1;
    if limits.nodes > MAX_NODES || depth > MAX_DEPTH {
        return Err(limit("XML structure"));
    }
    Ok(())
}

pub(crate) fn ownership(graph: &Graph) -> BTreeSet<String> {
    graph
        .charts
        .iter()
        .flat_map(|chart| {
            std::iter::once(chart.part_name.clone())
                .chain(
                    chart
                        .styles
                        .iter()
                        .map(|resource| resource.part_name.clone()),
                )
                .chain(
                    chart
                        .color_styles
                        .iter()
                        .map(|resource| resource.part_name.clone()),
                )
                .chain(
                    chart
                        .workbook
                        .iter()
                        .map(|resource| resource.part_name.clone()),
                )
        })
        .collect()
}

pub(crate) fn relationship_target(
    part: &dyn Part,
    relationship: &litchi_opc::Relationship,
) -> Result<PackURI> {
    if relationship.is_external() {
        return Err(invalid("external relationship is rejected"));
    }
    PackURI::from_rel_ref(part.partname().base_uri(), relationship.target_ref())
        .map_err(Error::InvalidFormat)
}

pub(crate) fn validate_leaf_path(uri: &PackURI, prefix: &str, label: &str) -> Result<()> {
    let Some(rest) = uri.as_str().strip_prefix(prefix) else {
        return Err(invalid(format!("{label} is outside {prefix}")));
    };
    if rest.is_empty()
        || rest.contains('/')
        || !rest
            .to_ascii_lowercase()
            .ends_with(if label == "embedded workbook" {
                ".xlsx"
            } else {
                ".xml"
            })
    {
        return Err(invalid(format!("invalid {label} path or suffix")));
    }
    Ok(())
}

pub(crate) fn require_content_type(part: &dyn Part, expected: &str, label: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(invalid(format!("{label} has invalid content type")))
    }
}

pub(crate) fn is_chart_rel(value: &str) -> bool {
    matches!(value, rt::CHART | rt::STRICT_CHART)
}

pub(crate) fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID is empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid("invalid relationship ID"))
    } else {
        Ok(())
    }
}

pub(crate) fn add_total(
    total: &mut usize,
    size: usize,
    individual: usize,
    label: &str,
) -> Result<()> {
    if size > individual {
        return Err(limit(label));
    }
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("aggregate bytes"))?;
    if *total > MAX_TOTAL_BYTES {
        return Err(limit("aggregate bytes"));
    }
    Ok(())
}

fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

pub(crate) fn limit(label: &str) -> Error {
    invalid(format!("DOCX chart {label} limit exceeded"))
}
