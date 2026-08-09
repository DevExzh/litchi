use super::super::model::{
    Chart, ChartSpaceFormatting, DataSet, Document, ExternalData, Info, PlotArea,
};
use super::limits::{
    CX, MAX_ATTRIBUTES, MAX_DATA_SETS, MAX_DEPTH, MAX_FEATURES, MAX_NODES, MAX_STRING_BYTES,
    MAX_XML_BYTES, R, R_STRICT,
};
use super::semantic::parse_data_graph;
use crate::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

#[derive(Debug)]
pub(super) struct Attribute {
    pub(super) namespace: String,
    pub(super) name: String,
    pub(super) value: String,
}

#[derive(Default)]
pub(super) struct Scan {
    root_depth: Option<usize>,
    root_closed: bool,
    root_rank: Option<u8>,
    root_children: HashSet<String>,
    chart_data_depth: Option<usize>,
    chart_depth: Option<usize>,
    data_depth: Option<usize>,
    current_data: Option<DataSet>,
    data_ids: HashSet<u32>,
    leaf_depth: Option<usize>,
    info: Option<Info>,
}

pub(super) fn parse_document(xml: &[u8]) -> Result<Document> {
    if xml.len() > MAX_XML_BYTES {
        return limit("ChartEx XML bytes");
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut scan = Scan::default();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_error(" depth overflow"))?;
                inspect_start(&reader, &element, depth, false, &mut scan, &mut strings)?;
            },
            Event::Empty(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_error(" depth overflow"))?;
                inspect_start(&reader, &element, depth, true, &mut scan, &mut strings)?;
                inspect_end(&reader, element.name(), depth, &mut scan)?;
                depth -= 1;
            },
            Event::End(element) => {
                inspect_end(&reader, element.name(), depth, &mut scan)?;
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("unbalanced  XML"))?;
            },
            Event::Text(value) => add_strings(&mut strings, value.as_ref().len())?,
            Event::CData(value) => add_strings(&mut strings, value.as_ref().len())?,
            Event::GeneralRef(value) => add_strings(&mut strings, value.as_ref().len())?,
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTD and processing instructions are rejected in  XML");
            },
            Event::Eof => break,
            _ => {},
        }
        nodes += 1;
        if nodes > MAX_NODES || depth > MAX_DEPTH {
            return limit(" XML structure");
        }
    }
    if depth != 0 || !scan.root_closed || scan.root_depth.is_none() {
        return invalid("missing or unterminated cx:chartSpace root");
    }
    let mut info = scan
        .info
        .ok_or_else(|| invalid_error("missing  metadata"))?;
    if info.data_sets.is_empty() {
        return invalid(" chartData requires at least one data set");
    }
    let (data_sets, series, axes, has_plot_surface, chart, plot_area, chart_space_formatting) =
        parse_data_graph(xml, &info.version, &info.features)?;
    info.data_sets = data_sets;
    info.series = series;
    info.axes = axes;
    info.has_plot_surface = has_plot_surface;
    info.chart = chart;
    info.plot_area = plot_area;
    info.chart_space_formatting = chart_space_formatting;
    Ok(Document {
        info,
        xml: xml.to_vec(),
        external_data_target: None,
        fallback_image_part_name: None,
        chart_style: None,
        chart_color_style: None,
    })
}

pub(super) fn inspect_start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    depth: usize,
    empty: bool,
    scan: &mut Scan,
    strings: &mut usize,
) -> Result<()> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let local_name = element.local_name();
    let local = std::str::from_utf8(local_name.as_ref()).map_err(xml_error)?;
    add_strings(strings, namespace.len() + local.len())?;
    let attributes = attributes(reader, element, strings)?;

    if depth == 1 {
        if scan.root_depth.is_some() || namespace != CX || local != "chartSpace" || empty {
            return invalid(" XML must have one non-empty cx:chartSpace root");
        }
        let version = optional(&attributes, "", "version")
            .unwrap_or("0.0")
            .to_owned();
        if version.len() > 64 {
            return limit(" version bytes");
        }
        let features = parse_features(optional(&attributes, "", "featureList").unwrap_or(""))?;
        let fallback = optional(&attributes, "", "fallbackImg")
            .filter(|value| !value.is_empty())
            .map(str::to_owned);
        if let Some(id) = &fallback {
            validate_id(id)?;
        }
        reject_unknown(
            &attributes,
            &[("", "version"), ("", "featureList"), ("", "fallbackImg")],
            "chartSpace",
        )?;
        scan.info = Some(Info {
            version,
            features,
            fallback_image_relationship_id: fallback,
            data_sets: Vec::new(),
            series: Vec::new(),
            axes: Vec::new(),
            has_plot_surface: false,
            chart: Chart::default(),
            plot_area: PlotArea::default(),
            chart_space_formatting: ChartSpaceFormatting::default(),
            external_data: None,
            has_title: false,
            has_legend: false,
        });
        scan.root_depth = Some(depth);
        return Ok(());
    }
    if scan.leaf_depth.is_some_and(|value| depth > value) {
        return invalid(" leaf element contains child elements");
    }
    let root_depth = scan
        .root_depth
        .ok_or_else(|| invalid_error("content before  root"))?;
    if depth == root_depth + 1 {
        let rank = root_child_rank(&namespace, local)
            .ok_or_else(|| invalid_error(format!("unsupported direct  child '{local}'")))?;
        if scan.root_rank.is_some_and(|previous| rank < previous)
            || !scan.root_children.insert(local.to_owned())
        {
            return invalid(" root children are duplicated or out of schema order");
        }
        scan.root_rank = Some(rank);
        match local {
            "chartData" => {
                if empty || rank != 0 {
                    return invalid(" chartData must be non-empty and first");
                }
                scan.chart_data_depth = Some(depth);
            },
            "chart" => {
                if empty || !scan.root_children.contains("chartData") {
                    return invalid(" chart must follow chartData");
                }
                scan.chart_depth = Some(depth);
            },
            _ => {},
        }
        return Ok(());
    }
    if scan
        .chart_data_depth
        .is_some_and(|value| depth == value + 1)
    {
        if namespace != CX {
            return invalid("foreign direct content in  chartData");
        }
        match local {
            "externalData" => {
                if scan
                    .info
                    .as_ref()
                    .is_some_and(|value| value.external_data.is_some())
                    || !empty
                {
                    return invalid(" externalData must be a unique leaf");
                }
                let id = required_any(&attributes, &[R, R_STRICT], "id")?.to_owned();
                validate_id(&id)?;
                let auto_update =
                    parse_bool(optional(&attributes, CX, "autoUpdate").unwrap_or("0"))?;
                scan.info
                    .as_mut()
                    .ok_or_else(|| invalid_error("chart root state is missing"))?
                    .external_data = Some(ExternalData {
                    relationship_id: id,
                    auto_update,
                });
            },
            "data" => {
                if empty || scan.current_data.is_some() {
                    return invalid(" data must be a non-empty direct chartData child");
                }
                let id = required(&attributes, "", "id")?
                    .parse::<u32>()
                    .map_err(|_err| invalid_error("invalid  data ID"))?;
                if !scan.data_ids.insert(id) || scan.data_ids.len() > MAX_DATA_SETS {
                    return invalid(" data IDs are duplicate or excessive");
                }
                scan.data_depth = Some(depth);
                scan.current_data = Some(DataSet {
                    id,
                    string_dimensions: 0,
                    numeric_dimensions: 0,
                    dimensions: Vec::new(),
                });
            },
            "extLst" => {},
            _ => return invalid("invalid direct  chartData child"),
        }
        return Ok(());
    }
    if scan.data_depth.is_some_and(|value| depth == value + 1) && namespace == CX {
        match local {
            "strDim" | "numDim" => {
                if empty || required(&attributes, "", "type")?.len() > 64 {
                    return invalid(" dimension requires a bounded type and content");
                }
                let data = scan
                    .current_data
                    .as_mut()
                    .ok_or_else(|| invalid_error("chart data state is missing"))?;
                if local == "strDim" {
                    data.string_dimensions += 1;
                } else {
                    data.numeric_dimensions += 1;
                }
            },
            "extLst" => {},
            _ => return invalid("invalid direct  data child"),
        }
        return Ok(());
    }
    if scan.chart_depth.is_some_and(|value| depth == value + 1) && namespace == CX {
        match local {
            "title" => {
                scan.info
                    .as_mut()
                    .ok_or_else(|| invalid_error("chart root state is missing"))?
                    .has_title = true;
            },
            "legend" => {
                scan.info
                    .as_mut()
                    .ok_or_else(|| invalid_error("chart root state is missing"))?
                    .has_legend = true;
            },
            "plotArea" | "extLst" => {},
            _ => return invalid("invalid direct  chart child"),
        }
    }
    Ok(())
}

pub(super) fn inspect_end(
    reader: &NsReader<&[u8]>,
    name: quick_xml::name::QName<'_>,
    depth: usize,
    scan: &mut Scan,
) -> Result<()> {
    let namespace = resolved(reader.resolver().resolve_element(name).0)?;
    let local_name = name.local_name();
    let local = std::str::from_utf8(local_name.as_ref()).map_err(xml_error)?;
    if scan.data_depth == Some(depth) && namespace == CX && local == "data" {
        let data = scan
            .current_data
            .take()
            .ok_or_else(|| invalid_error("chart data state is missing"))?;
        if data.string_dimensions + data.numeric_dimensions == 0 {
            return invalid(" data requires at least one dimension");
        }
        scan.info
            .as_mut()
            .ok_or_else(|| invalid_error("chart root state is missing"))?
            .data_sets
            .push(data);
        scan.data_depth = None;
    } else if scan.chart_data_depth == Some(depth) && namespace == CX && local == "chartData" {
        scan.chart_data_depth = None;
    } else if scan.chart_depth == Some(depth) && namespace == CX && local == "chart" {
        scan.chart_depth = None;
    } else if scan.root_depth == Some(depth) && namespace == CX && local == "chartSpace" {
        if !scan.root_children.contains("chartData") || !scan.root_children.contains("chart") {
            return invalid(" root requires chartData followed by chart");
        }
        scan.root_closed = true;
    }
    if scan.leaf_depth == Some(depth) {
        scan.leaf_depth = None;
    }
    Ok(())
}

pub(super) fn root_child_rank(namespace: &str, local: &str) -> Option<u8> {
    if namespace != CX {
        return None;
    }
    match local {
        "chartData" => Some(0),
        "chart" => Some(1),
        "spPr" => Some(2),
        "txPr" => Some(3),
        "clrMapOvr" => Some(4),
        "fmtOvrs" => Some(5),
        "printSettings" => Some(6),
        "extLst" => Some(7),
        _ => None,
    }
}

#[derive(Debug)]
pub(super) struct MiniNode {
    pub(super) namespace: String,
    pub(super) name: String,
    pub(super) attributes: Vec<Attribute>,
    pub(super) children: Vec<MiniNode>,
    pub(super) text: String,
}

pub(super) fn parse_mini_tree(xml: &[u8]) -> Result<MiniNode> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::<MiniNode>::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return limit(" semantic XML structure");
                }
                let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
                let local_name = element.local_name();
                let name = std::str::from_utf8(local_name.as_ref())
                    .map_err(xml_error)?
                    .to_owned();
                add_strings(&mut strings, namespace.len() + name.len())?;
                let node = MiniNode {
                    namespace,
                    name,
                    attributes: attributes(&reader, &element, &mut strings)?,
                    children: Vec::new(),
                    text: String::new(),
                };
                stack.push(node);
            },
            Event::Empty(element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return limit(" semantic XML structure");
                }
                let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
                let local_name = element.local_name();
                let name = std::str::from_utf8(local_name.as_ref())
                    .map_err(xml_error)?
                    .to_owned();
                add_strings(&mut strings, namespace.len() + name.len())?;
                let node = MiniNode {
                    namespace,
                    name,
                    attributes: attributes(&reader, &element, &mut strings)?,
                    children: Vec::new(),
                    text: String::new(),
                };
                attach_mini(node, &mut stack, &mut root)?;
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid_error("unexpected  closing element"))?;
                attach_mini(node, &mut stack, &mut root)?;
            },
            Event::Text(value) => {
                let decoded = value.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return invalid("text outside  root");
                }
            },
            Event::CData(value) => {
                let decoded = value.decode().map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = match name.as_ref() {
                    "amp" => "&",
                    "lt" => "<",
                    "gt" => ">",
                    "apos" => "'",
                    "quot" => "\"",
                    _ => return invalid("custom entity in  data is rejected"),
                };
                add_strings(&mut strings, value.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(value);
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTD and processing instructions are rejected in  data");
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return invalid("unterminated  semantic XML");
    }
    root.ok_or_else(|| invalid_error("missing  semantic root"))
}

pub(super) fn attach_mini(
    node: MiniNode,
    stack: &mut [MiniNode],
    root: &mut Option<MiniNode>,
) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return invalid("multiple  XML roots");
    }
    Ok(())
}

pub(super) fn one_child<'a>(
    node: &'a MiniNode,
    namespace: &str,
    name: &str,
) -> Result<Option<&'a MiniNode>> {
    let mut values = node
        .children
        .iter()
        .filter(|value| value.namespace == namespace && value.name == name);
    let value = values.next();
    if values.next().is_some() {
        invalid(format!("duplicate  {name}"))
    } else {
        Ok(value)
    }
}

pub(super) fn parse_u32(value: &str, label: &str) -> Result<u32> {
    value
        .parse()
        .map_err(|_err| invalid_error(format!("invalid  {label}")))
}
pub(super) fn parse_i32(value: &str, label: &str) -> Result<i32> {
    value
        .parse()
        .map_err(|_err| invalid_error(format!("invalid  {label}")))
}
pub(super) fn bounded_optional(node: &MiniNode, name: &str, max: usize) -> Result<Option<String>> {
    optional(&node.attributes, "", name)
        .map(|value| {
            if value.len() <= max {
                Ok(value.to_owned())
            } else {
                limit(" attribute string")
            }
        })
        .transpose()
}
pub(super) fn valid_xml_double(value: &str) -> bool {
    matches!(value, "INF" | "-INF" | "NaN") || (!value.is_empty() && value.parse::<f64>().is_ok())
}

pub(super) fn attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    strings: &mut usize,
) -> Result<Vec<Attribute>> {
    let mut values = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        if item.key.as_ref() == b"xmlns" || item.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if values.len() >= MAX_ATTRIBUTES {
            return limit(" element attributes");
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        add_strings(strings, namespace.len() + name.len() + value.len())?;
        if values
            .iter()
            .any(|existing: &Attribute| existing.namespace == namespace && existing.name == name)
        {
            return invalid("duplicate expanded  attribute");
        }
        values.push(Attribute {
            namespace,
            name,
            value,
        });
    }
    Ok(values)
}

pub(super) fn parse_features(value: &str) -> Result<Vec<String>> {
    let features = value
        .split_ascii_whitespace()
        .map(str::to_owned)
        .collect::<Vec<_>>();
    if features.len() > MAX_FEATURES || features.iter().any(|value| value.len() > 128) {
        return limit(" feature list");
    }
    Ok(features)
}

pub(super) fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "0" | "false" => Ok(false),
        "1" | "true" => Ok(true),
        _ => invalid("invalid  boolean"),
    }
}

pub(super) fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return invalid(" relationship ID is empty");
    };
    if value.len() > 255
        || !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        return invalid("invalid  relationship ID");
    }
    Ok(())
}

pub(super) fn optional<'a>(
    attributes: &'a [Attribute],
    namespace: &str,
    name: &str,
) -> Option<&'a str> {
    attributes
        .iter()
        .find(|value| value.namespace == namespace && value.name == name)
        .map(|value| value.value.as_str())
}
pub(super) fn required<'a>(
    attributes: &'a [Attribute],
    namespace: &str,
    name: &str,
) -> Result<&'a str> {
    optional(attributes, namespace, name)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid_error(format!("missing  attribute '{name}'")))
}
pub(super) fn required_any<'a>(
    attributes: &'a [Attribute],
    namespaces: &[&str],
    name: &str,
) -> Result<&'a str> {
    let mut values = attributes
        .iter()
        .filter(|value| namespaces.contains(&value.namespace.as_str()) && value.name == name);
    let value = values
        .next()
        .ok_or_else(|| invalid_error(format!("missing  relationship attribute '{name}'")))?;
    if values.next().is_some() {
        return invalid("duplicate  relationship attribute aliases");
    }
    Ok(&value.value)
}
pub(super) fn reject_unknown(
    attributes: &[Attribute],
    allowed: &[(&str, &str)],
    element: &str,
) -> Result<()> {
    if attributes
        .iter()
        .any(|value| !allowed.contains(&(value.namespace.as_str(), value.name.as_str())))
    {
        return invalid(format!("unexpected attribute on  {element}"));
    }
    Ok(())
}
pub(super) fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound  prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}
pub(super) fn add_strings(total: &mut usize, amount: usize) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| invalid_error(" string size overflow"))?;
    if *total > MAX_STRING_BYTES {
        limit(" string bytes")
    } else {
        Ok(())
    }
}
pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Invalid(format!("invalid  XML: {error}"))
}
pub(super) fn invalid_error(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
pub(super) fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}
pub(super) fn limit<T>(name: &str) -> Result<T> {
    invalid(format!("{name} exceeds resource limit"))
}
