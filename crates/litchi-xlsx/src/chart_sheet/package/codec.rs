//! Bounded XML and resource codecs for the chartsheet package graph.

use super::super::{Conformance, State};
use super::{
    CHART_EX, CHART_EX_CHOICE, DRAWING_MAIN, MAX_CHART_BYTES, MAX_CHART_DIRECT_IMAGES,
    MAX_CHART_EX_BYTES, MAX_CHART_THEME_IMAGES, MAX_CHART_THEME_OVERRIDE_BYTES,
    MAX_CHART_USER_SHAPE_IMAGES, MAX_CHART_USER_SHAPES_BYTES, MAX_CHARTS, MAX_DEPTH,
    MAX_DRAWING_BYTES, MAX_NAMESPACE_BINDINGS, MAX_NODES, MAX_STRING_BYTES,
    MAX_TOTAL_RESOURCE_BYTES, SML, STRICT_DRAWING_MAIN, STRICT_SML, invalid, limit,
};
use crate::{Error, Result};
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet, HashSet};

#[derive(Clone)]
pub(super) struct Attribute {
    pub(super) namespace: String,
    pub(super) name: String,
    pub(super) value: String,
}

#[derive(Clone)]
pub(super) struct Node {
    pub(super) namespace: String,
    pub(super) name: String,
    pub(super) attributes: Vec<Attribute>,
    pub(super) children: Vec<Node>,
    pub(super) text: String,
    content: Vec<NodeContent>,
}

#[derive(Clone)]
enum NodeContent {
    Text(String),
    Child,
}

pub(super) fn collect_extension_relationship_ids(
    node: &Node,
    relationship_namespace: &str,
    ids: &mut BTreeSet<String>,
    max_ids: usize,
) -> Result<()> {
    for attribute in &node.attributes {
        if attribute.namespace == relationship_namespace {
            validate_id(&attribute.value)?;
            if !ids.contains(&attribute.value) && ids.len() >= max_ids {
                return Err(limit("relationship reference count"));
            }
            ids.insert(attribute.value.clone());
        }
    }
    for child in &node.children {
        collect_extension_relationship_ids(child, relationship_namespace, ids, max_ids)?;
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum DrawingChartKind {
    Classic,
    Extended,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct DrawingChartReference {
    pub(super) relationship_id: String,
    pub(super) kind: DrawingChartKind,
}

pub(super) fn chart_ex_mce_capabilities() -> Capabilities {
    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities
        .understand_namespace(CHART_EX)
        .understand_namespace(CHART_EX_CHOICE);
    capabilities
}
pub(super) fn drawing_chart_references(
    xml: &[u8],
    conformance: Conformance,
) -> Result<Vec<DrawingChartReference>> {
    if xml.len() > MAX_DRAWING_BYTES {
        return Err(limit("drawing bytes"));
    }
    let root =
        parse_document_with_capabilities(xml, MAX_DRAWING_BYTES, &chart_ex_mce_capabilities())?;
    if root.namespace != conformance.xdr() || root.name != "wsDr" {
        return Err(invalid(
            "drawing root does not match chartsheet conformance",
        ));
    }
    let mut references = Vec::new();
    collect_drawing_chart_references(&root, conformance, &mut references)?;
    if references.len() > MAX_CHARTS {
        return Err(limit("chart count"));
    }
    let mut ids = HashSet::new();
    for reference in &references {
        validate_id(&reference.relationship_id)?;
        if !ids.insert(reference.relationship_id.as_str()) {
            return Err(invalid("drawing chart relationship IDs collide"));
        }
    }
    Ok(references)
}
pub(super) fn collect_drawing_chart_references(
    node: &Node,
    conformance: Conformance,
    references: &mut Vec<DrawingChartReference>,
) -> Result<()> {
    if matches!(node.namespace.as_str(), DRAWING_MAIN | STRICT_DRAWING_MAIN)
        && node.name == "graphicData"
    {
        let has_chart_ex = node
            .children
            .iter()
            .any(|child| child.namespace == CHART_EX && child.name == "chart");
        if has_chart_ex || optional(node, "", "uri") == Some(CHART_EX) {
            if optional(node, "", "uri") != Some(CHART_EX) {
                return Err(invalid(
                    "cx:chart requires the exact chartEx graphicData URI",
                ));
            }
            whitespace(node)?;
            no_attributes(node, &[("", "uri")])?;
            if node.children.len() != 1 {
                return Err(invalid(
                    "chartEx graphicData requires exactly one cx:chart child",
                ));
            }
            let chart = node
                .children
                .first()
                .ok_or_else(|| invalid("chartEx graphicData is missing its chart"))?;
            if chart.namespace != CHART_EX || chart.name != "chart" {
                return Err(invalid("chartEx graphicData has an invalid root child"));
            }
            leaf(chart, "chartEx drawing reference")?;
            whitespace(chart)?;
            no_attributes(chart, &[(conformance.rel(), "id")])?;
            if references.len() >= MAX_CHARTS {
                return Err(limit("chart count"));
            }
            references.push(DrawingChartReference {
                relationship_id: required(chart, conformance.rel(), "id")?.to_owned(),
                kind: DrawingChartKind::Extended,
            });
            return Ok(());
        }
    }
    if node.namespace == conformance.chart() && node.name == "chart" {
        if references.len() >= MAX_CHARTS {
            return Err(limit("chart count"));
        }
        references.push(DrawingChartReference {
            relationship_id: required(node, conformance.rel(), "id")?.to_owned(),
            kind: DrawingChartKind::Classic,
        });
    }
    for child in &node.children {
        collect_drawing_chart_references(child, conformance, references)?;
    }
    Ok(())
}

pub(super) fn validate_chart_xml(xml: &[u8], conformance: Conformance) -> Result<()> {
    if xml.len() > MAX_CHART_BYTES {
        return Err(limit("chart bytes"));
    }
    let root = parse_document(xml, MAX_CHART_BYTES)?;
    if root.namespace == conformance.chart() && root.name == "chartSpace" {
        Ok(())
    } else {
        Err(invalid("chart root does not match chartsheet conformance"))
    }
}
pub(super) fn validate_chart_companion_xml(
    xml: &[u8],
    root_name: &str,
    max_bytes: usize,
) -> Result<()> {
    if xml.len() > max_bytes {
        return Err(limit("chart companion bytes"));
    }
    let result = match root_name {
        "chartStyle" => litchi_drawingml::chart::style::parse(xml).map(|_| ()),
        "colorStyle" => litchi_drawingml::chart::style::parse_color(xml).map(|_| ()),
        _ => {
            return Err(invalid(format!(
                "unsupported chart companion root '{root_name}'"
            )));
        },
    };
    result.map_err(Error::from)
}
pub(super) fn validate_chart_user_shapes_xml(
    xml: &[u8],
    conformance: Conformance,
) -> Result<BTreeSet<String>> {
    if xml.len() > MAX_CHART_USER_SHAPES_BYTES {
        return Err(limit("chartUserShapes bytes"));
    }
    let root = parse_document(xml, MAX_CHART_USER_SHAPES_BYTES)?;
    if root.namespace != conformance.chart() || root.name != "userShapes" {
        return Err(invalid(
            "chartUserShapes root does not match chartsheet conformance",
        ));
    }
    let mut ids = BTreeSet::new();
    collect_extension_relationship_ids(
        &root,
        conformance.rel(),
        &mut ids,
        MAX_CHART_USER_SHAPE_IMAGES,
    )?;
    Ok(ids)
}
#[derive(Default)]
pub(super) struct ChartExRelationshipReferences {
    pub(super) images: BTreeSet<String>,
    pub(super) package: Option<String>,
}
pub(super) fn validate_chart_ex_relationships(
    xml: &[u8],
    conformance: Conformance,
) -> Result<ChartExRelationshipReferences> {
    if xml.len() > MAX_CHART_EX_BYTES {
        return Err(limit("chartEx bytes"));
    }
    let root = parse_document(xml, MAX_CHART_EX_BYTES)?;
    if root.namespace != CHART_EX || root.name != "chartSpace" {
        return Err(invalid("invalid chartEx root"));
    }
    let mut references = ChartExRelationshipReferences::default();
    collect_chart_ex_relationships(&root, conformance, &mut references)?;
    Ok(references)
}
pub(super) fn collect_chart_ex_relationships(
    node: &Node,
    conformance: Conformance,
    references: &mut ChartExRelationshipReferences,
) -> Result<()> {
    let external_data = node.namespace == CHART_EX && node.name == "externalData";
    if external_data
        && optional(node, CHART_EX, "autoUpdate").is_some_and(|value| matches!(value, "1" | "true"))
    {
        return Err(invalid("auto-updating chartEx external data is rejected"));
    }
    for attribute in &node.attributes {
        if attribute.namespace == conformance.rel() {
            validate_id(&attribute.value)?;
            if external_data && attribute.name == "id" {
                if references
                    .package
                    .replace(attribute.value.clone())
                    .is_some()
                {
                    return Err(invalid(
                        "chartEx has multiple externalData package references",
                    ));
                }
            } else {
                if external_data {
                    return Err(invalid(
                        "chartEx externalData has an unsupported relationship attribute",
                    ));
                }
                if !references.images.contains(&attribute.value)
                    && references.images.len() >= MAX_CHART_DIRECT_IMAGES
                {
                    return Err(limit("chartEx direct image reference count"));
                }
                references.images.insert(attribute.value.clone());
            }
        }
    }
    for child in &node.children {
        collect_chart_ex_relationships(child, conformance, references)?;
    }
    Ok(())
}
pub(super) fn validate_theme_override_xml(
    xml: &[u8],
    conformance: Conformance,
) -> Result<BTreeSet<String>> {
    if xml.len() > MAX_CHART_THEME_OVERRIDE_BYTES {
        return Err(limit("themeOverride bytes"));
    }
    let root = parse_document(xml, MAX_CHART_THEME_OVERRIDE_BYTES)?;
    let namespace = if conformance == Conformance::Strict {
        STRICT_DRAWING_MAIN
    } else {
        DRAWING_MAIN
    };
    if root.namespace != namespace || root.name != "themeOverride" {
        return Err(invalid(
            "themeOverride root does not match chartsheet conformance",
        ));
    }
    let mut ids = BTreeSet::new();
    collect_extension_relationship_ids(&root, conformance.rel(), &mut ids, MAX_CHART_THEME_IMAGES)?;
    Ok(ids)
}

pub(super) fn parse_document(xml: &[u8], max_bytes: usize) -> Result<Node> {
    parse_document_with_capabilities(xml, max_bytes, &Capabilities::ooxml_baseline())
}

pub(super) fn parse_document_with_capabilities(
    xml: &[u8],
    max_bytes: usize,
    capabilities: &Capabilities,
) -> Result<Node> {
    if xml.len() > max_bytes {
        return Err(limit("input XML bytes"));
    }
    let limits = Limits {
        max_input_bytes: max_bytes,
        max_output_bytes: max_bytes,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: MAX_NAMESPACE_BINDINGS,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, capabilities, &limits)?.xml;
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("XML node count"))?;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                let node = make_node(&reader, element, reader.decoder(), &mut strings)?;
                if empty {
                    attach(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    add_node_text(node, &decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|v| v.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| invalid("custom XML entity is rejected"))?;
                add_strings(&mut strings, value.len())?;
                if let Some(node) = stack.last_mut() {
                    add_node_text(node, &value);
                }
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        };
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    root.ok_or_else(|| invalid("missing XML root"))
}

pub(super) fn make_node(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    strings: &mut usize,
) -> Result<Node> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    add_strings(strings, namespace.len() + name.len())?;
    let mut attributes = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let qname = item.key.as_ref();
        if qname == b"xmlns" || qname.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_strings(strings, namespace.len() + name.len() + value.len())?;
        if attributes
            .iter()
            .any(|a: &Attribute| a.namespace == namespace && a.name == name)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        attributes.push(Attribute {
            namespace,
            name,
            value,
        });
    }
    Ok(Node {
        namespace,
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
        content: Vec::new(),
    })
}

pub(super) fn add_node_text(node: &mut Node, value: &str) {
    node.text.push_str(value);
    match node.content.last_mut() {
        Some(NodeContent::Text(current)) => current.push_str(value),
        _ => node.content.push(NodeContent::Text(value.to_owned())),
    }
}
pub(super) fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
        parent.content.push(NodeContent::Child);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}
pub(super) fn root_conformance(root: &Node, name: &str) -> Result<Conformance> {
    if root.name != name {
        return Err(invalid(format!("expected {name} root")));
    }
    match root.namespace.as_str() {
        SML => Ok(Conformance::Transitional),
        STRICT_SML => Ok(Conformance::Strict),
        _ => Err(invalid("unsupported SpreadsheetML namespace")),
    }
}
pub(super) fn one_child<'a>(
    node: &'a Node,
    namespace: &str,
    name: &str,
) -> Result<Option<&'a Node>> {
    let mut values = node
        .children
        .iter()
        .filter(|c| c.namespace == namespace && c.name == name);
    let value = values.next();
    if values.next().is_some() {
        Err(invalid(format!(
            "{} has multiple {name} children",
            node.name
        )))
    } else {
        Ok(value)
    }
}
pub(super) fn required_child<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a Node> {
    one_child(node, namespace, name)?
        .ok_or_else(|| invalid(format!("{} is missing {name}", node.name)))
}
pub(super) fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|a| a.namespace == namespace && a.name == name)
        .map(|a| a.value.as_str())
}
pub(super) fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> {
    optional(node, namespace, name)
        .filter(|v| !v.is_empty())
        .ok_or_else(|| invalid(format!("{} is missing attribute '{name}'", node.name)))
}
pub(super) fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    if let Some(attribute) = node
        .attributes
        .iter()
        .find(|a| !allowed.contains(&(a.namespace.as_str(), a.name.as_str())))
    {
        Err(invalid(format!(
            "unexpected attribute '{}' on {}",
            attribute.name, node.name
        )))
    } else {
        Ok(())
    }
}
pub(super) fn whitespace(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}
pub(super) fn leaf(node: &Node, label: &str) -> Result<()> {
    if node.children.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{label} must not contain child elements")))
    }
}
pub(super) fn parse_state(value: &str) -> Result<State> {
    match value {
        "visible" => Ok(State::Visible),
        "hidden" => Ok(State::Hidden),
        "veryHidden" => Ok(State::VeryHidden),
        _ => Err(invalid("invalid workbook sheet state")),
    }
}
pub(super) fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID cannot be empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|b| b.is_ascii_alphanumeric() || matches!(b, b'_' | b'-' | b'.'))
    {
        Err(invalid(format!("invalid relationship ID '{value}'")))
    } else {
        Ok(())
    }
}
pub(super) fn bounded(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("string bytes"))
    }
}
pub(super) fn add_strings(total: &mut usize, size: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
}
pub(super) fn resolved(value: ResolveResult<'_>) -> Result<String> {
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
pub(super) fn internal_relationship<'a>(
    part: &'a dyn Part,
    id: &str,
    kind: &str,
) -> Result<&'a litchi_opc::Relationship> {
    let relationship = part
        .rels()
        .get(id)
        .ok_or_else(|| invalid(format!("missing relationship '{id}'")))?;
    if relationship.reltype() != kind {
        return Err(invalid(format!("relationship '{id}' has unexpected type")));
    }
    if relationship.is_external() {
        return Err(invalid(format!(
            "external relationship '{id}' is not loaded"
        )));
    }
    Ok(relationship)
}
pub(super) fn require_workbook(part: &dyn Part) -> Result<()> {
    if matches!(
        part.content_type(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            | "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            | "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
            | "application/vnd.ms-excel.template.macroEnabled.main+xml"
    ) {
        Ok(())
    } else {
        Err(invalid("source part is not a workbook"))
    }
}
pub(super) fn require_content_type(part: &dyn Part, expected: &str, label: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(invalid(format!(
            "{label} part has content type '{}'",
            part.content_type()
        )))
    }
}
pub(super) fn new_uri(package: &OpcPackage, value: &str, prefix: &str) -> Result<PackURI> {
    let uri = PackURI::new(value).map_err(invalid)?;
    if !uri.as_str().starts_with(prefix) {
        return Err(invalid(format!("part '{uri}' is outside {prefix}")));
    }
    package.validate_new_part_name(&uri)?;
    Ok(uri)
}
pub(super) fn staged_uri<K: Ord>(
    uris: &BTreeMap<K, PackURI>,
    key: &K,
    label: &str,
) -> Result<PackURI> {
    uris.get(key)
        .cloned()
        .ok_or_else(|| invalid(format!("missing staged {label} URI")))
}
pub(super) fn add_relationship_checked(
    package: &mut OpcPackage,
    source: &PackURI,
    relationship_type: &str,
    target: String,
    relationship_id: String,
    target_mode: TargetMode,
) -> Result<()> {
    package
        .get_part_mut(source)?
        .rels_mut()
        .try_add_relationship(
            relationship_type.to_owned(),
            target,
            relationship_id,
            target_mode,
        )?;
    Ok(())
}
pub(super) fn add_resource(
    total: &mut usize,
    size: usize,
    individual: usize,
    name: &str,
) -> Result<()> {
    if size > individual {
        return Err(limit(name));
    }
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("total resource bytes"))?;
    if *total > MAX_TOTAL_RESOURCE_BYTES {
        Err(limit("total resource bytes"))
    } else {
        Ok(())
    }
}

pub(super) fn attr(out: &mut Vec<u8>, name: &str, value: &str) {
    out.push(b' ');
    out.extend_from_slice(name.as_bytes());
    out.extend_from_slice(b"=\"");
    escape(out, value);
    out.push(b'\"');
}
pub(super) fn escape(out: &mut Vec<u8>, value: &str) {
    for c in value.chars() {
        match c {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '"' => out.extend_from_slice(b"&quot;"),
            '\t' => out.extend_from_slice(b"&#x9;"),
            '\n' => out.extend_from_slice(b"&#xA;"),
            '\r' => out.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                out.extend_from_slice(c.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
