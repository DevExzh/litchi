//! Bounded strict/transitional `SpreadsheetML` XML and opaque-fragment codec.

use super::model::{
    CacheDefinition, CachePivotTable, FilterType, Level, OpaqueXml, PivotFilter, Range, State,
    View, Views, validate_cache_definition, validate_range, validate_state,
    validate_timeline_filter, validate_views_local,
};
use super::{
    MAX_DEPTH, MAX_NODES, MAX_OPAQUE_BYTES, MAX_PIVOT_TABLES, MAX_STRING_BYTES, MAX_TIMELINES,
    MAX_XML_BYTES, REL, SML, STRICT_REL, STRICT_SML, X15, XR10, bounded, invalid, limit, xml_error,
};
use crate::auto_filter::{parse_auto_filter_fragment, write_auto_filter_fragment};
use crate::error::Result;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, HashMap, HashSet};

fn parse_range(n: &Node) -> Result<Range> {
    no_attributes(n, &[("", "startDate"), ("", "endDate")])?;
    empty(n)?;
    Range::new(required(n, "", "startDate")?, required(n, "", "endDate")?)
}
fn parse_state(n: &Node) -> Result<State> {
    require(n, X15, "state")?;
    whitespace(n)?;
    no_attributes(
        n,
        &[
            ("", "singleRangeFilterState"),
            ("", "minimalRefreshVersion"),
            ("", "lastRefreshVersion"),
            ("", "pivotCacheId"),
            ("", "filterType"),
        ],
    )?;
    let mut selection = None;
    let mut bounds = None;
    let mut extension_list = None;
    let mut stage = 0;
    for c in &n.children {
        match (c.namespace.as_str(), c.name.as_str()) {
            (X15, "selection") if stage == 0 => {
                selection = Some(parse_range(c)?);
                stage = 1;
            },
            (X15, "bounds") if stage <= 1 && bounds.is_none() => {
                bounds = Some(parse_range(c)?);
                stage = 2;
            },
            (ns, "extLst") if stage == 2 && (ns == X15 || ns == SML || ns == STRICT_SML) => {
                extension_list = Some(opaque_from_node(c)?);
                stage = 3;
            },
            _ => return Err(invalid("invalid or out-of-order View state child")),
        }
    }
    let state = State {
        selection,
        bounds: bounds.ok_or_else(|| invalid("View state requires bounds"))?,
        extension_list,
        single_range_filter_state: optional_bool(n, "singleRangeFilterState")?,
        minimal_refresh_version: required(n, "", "minimalRefreshVersion")?
            .parse()
            .map_err(|_source| invalid("invalid minimalRefreshVersion"))?,
        last_refresh_version: required(n, "", "lastRefreshVersion")?
            .parse()
            .map_err(|_source| invalid("invalid lastRefreshVersion"))?,
        pivot_cache_id: required(n, "", "pivotCacheId")?
            .parse()
            .map_err(|_source| invalid("invalid pivotCacheId"))?,
        filter_type: FilterType::parse(required(n, "", "filterType")?)?,
    };
    validate_state(&state)?;
    Ok(state)
}
fn parse_timeline_pivot_filter(n: &Node) -> Result<PivotFilter> {
    require(n, X15, "timelinePivotFilter")?;
    whitespace(n)?;
    no_attributes(
        n,
        &[
            ("", "useWholeDay"),
            ("", "fld"),
            ("", "id"),
            ("", "name"),
            ("", "description"),
        ],
    )?;
    if n.children.len() > 1 {
        return Err(invalid(
            "timelinePivotFilter permits at most one autoFilter",
        ));
    }
    let auto_filter = n
        .children
        .first()
        .map(|c| {
            if !matches!(c.namespace.as_str(), SML | STRICT_SML) || c.name != "autoFilter" {
                return Err(invalid("timelinePivotFilter child must be autoFilter"));
            }
            parse_auto_filter_fragment(&serialize_node(c)?)
        })
        .transpose()?;
    let v = PivotFilter {
        use_whole_day: optional_bool(n, "useWholeDay")?,
        field: required(n, "", "fld")?
            .parse()
            .map_err(|_source| invalid("invalid timeline filter fld"))?,
        id: required(n, "", "id")?
            .parse()
            .map_err(|_source| invalid("invalid timeline filter id"))?,
        name: optional(n, "", "name").map(str::to_owned),
        description: optional(n, "", "description").map(str::to_owned),
        auto_filter,
    };
    validate_timeline_filter(&v)?;
    Ok(v)
}
fn write_range(x: &mut Vec<u8>, name: &str, v: &Range) -> Result<()> {
    validate_range(v)?;
    x.extend_from_slice(b"<x15:");
    x.extend_from_slice(name.as_bytes());
    attr(x, "startDate", &v.start_date);
    attr(x, "endDate", &v.end_date);
    x.extend_from_slice(b"/>");
    Ok(())
}
fn write_state(x: &mut Vec<u8>, v: &State) -> Result<()> {
    validate_state(v)?;
    x.extend_from_slice(b"<x15:state");
    if let Some(q) = v.single_range_filter_state {
        bool_attr(x, "singleRangeFilterState", q);
    }
    attr(
        x,
        "minimalRefreshVersion",
        &v.minimal_refresh_version.to_string(),
    );
    attr(x, "lastRefreshVersion", &v.last_refresh_version.to_string());
    attr(x, "pivotCacheId", &v.pivot_cache_id.to_string());
    attr(x, "filterType", v.filter_type.as_str());
    x.push(b'>');
    if let Some(q) = &v.selection {
        write_range(x, "selection", q)?;
    }
    write_range(x, "bounds", &v.bounds)?;
    if let Some(q) = &v.extension_list {
        append_opaque_any_namespace(x, q, "extLst")?;
    }
    x.extend_from_slice(b"</x15:state>");
    Ok(())
}
fn write_timeline_pivot_filter(x: &mut Vec<u8>, v: &PivotFilter) -> Result<()> {
    validate_timeline_filter(v)?;
    x.extend_from_slice(b"<x15:timelinePivotFilter");
    if let Some(q) = v.use_whole_day {
        bool_attr(x, "useWholeDay", q);
    }
    attr(x, "fld", &v.field.to_string());
    attr(x, "id", &v.id.to_string());
    if let Some(q) = &v.name {
        attr(x, "name", q);
    }
    if let Some(q) = &v.description {
        attr(x, "description", q);
    }
    if let Some(q) = &v.auto_filter {
        x.push(b'>');
        x.extend_from_slice(&write_auto_filter_fragment(q)?);
        x.extend_from_slice(b"</x15:timelinePivotFilter>");
    } else {
        x.extend_from_slice(b"/>");
    }
    Ok(())
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Attribute {
    pub(super) namespace: String,
    pub(super) name: String,
    pub(super) value: String,
}
#[derive(Clone, Debug)]
pub(super) struct Node {
    pub(super) namespace: String,
    pub(super) name: String,
    pub(super) attributes: Vec<Attribute>,
    pub(super) children: Vec<Node>,
    pub(super) text: String,
}

/// Parse one MS-XLSX View Cache part.
pub fn parse_timeline_cache_definition(xml: &[u8]) -> Result<CacheDefinition> {
    let root = parse_document(xml)?;
    require(&root, X15, "timelineCacheDefinition")?;
    whitespace(&root)?;
    no_attributes(&root, &[("", "name"), ("", "sourceName"), (XR10, "uid")])?;
    let name = required(&root, "", "name")?.to_owned();
    let source_name = required(&root, "", "sourceName")?.to_owned();
    let uid = optional(&root, XR10, "uid").map(str::to_owned);
    let mut pivot_tables = Vec::new();
    let mut state = None;
    let mut timeline_pivot_filter = None;
    let mut extension_list = None;
    let mut stage = 0u8;
    for child in &root.children {
        match child.name.as_str() {
            "pivotTables" if child.namespace == X15 && stage == 0 => {
                stage = 1;
                parse_pivot_tables(child, &mut pivot_tables)?;
            },
            "state" if child.namespace == X15 && stage <= 1 && state.is_none() => {
                stage = 2;
                state = Some(parse_state(child)?);
            },
            "timelinePivotFilter"
                if child.namespace == X15 && stage == 2 && timeline_pivot_filter.is_none() =>
            {
                stage = 3;
                timeline_pivot_filter = Some(parse_timeline_pivot_filter(child)?);
            },
            "extLst"
                if (child.namespace == X15
                    || child.namespace == SML
                    || child.namespace == STRICT_SML)
                    && stage >= 2
                    && extension_list.is_none() =>
            {
                stage = 4;
                extension_list = Some(opaque_from_node(child)?);
            },
            _ => {
                return Err(invalid(format!(
                    "unexpected or out-of-order View Cache child '{}'",
                    child.name
                )));
            },
        }
    }
    let value = CacheDefinition {
        name,
        uid,
        source_name,
        pivot_tables,
        state: state.ok_or_else(|| invalid("View Cache requires exactly one state element"))?,
        timeline_pivot_filter,
        extension_list,
    };
    validate_cache_definition(&value)?;
    Ok(value)
}

fn parse_pivot_tables(node: &Node, output: &mut Vec<CachePivotTable>) -> Result<()> {
    no_attributes(node, &[])?;
    whitespace(node)?;
    if node.children.is_empty() {
        return Err(invalid("pivotTables must contain at least one pivotTable"));
    }
    if node.children.len() > MAX_PIVOT_TABLES {
        return Err(limit("pivot table count"));
    }
    let mut bindings = HashSet::new();
    for child in &node.children {
        require(child, X15, "pivotTable")?;
        no_attributes(child, &[("", "tabId"), ("", "name")])?;
        empty(child)?;
        let tab_id = required(child, "", "tabId")?
            .parse::<u32>()
            .map_err(|_source| invalid("invalid View Cache pivotTable tabId"))?;
        let name = required(child, "", "name")?.to_owned();
        bounded(&name, "pivot table name")?;
        if !bindings.insert((tab_id, name.to_lowercase())) {
            return Err(invalid("duplicate View Cache pivotTable binding"));
        }
        output.push(CachePivotTable { tab_id, name });
    }
    Ok(())
}

/// Deterministically serialize one View Cache part.
pub fn write_timeline_cache_definition(value: &CacheDefinition) -> Result<Vec<u8>> {
    validate_cache_definition(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<x15:timelineCacheDefinition xmlns:x15=\"");
    escape(&mut output, X15);
    output.extend_from_slice(b"\" xmlns:xr10=\"");
    escape(&mut output, XR10);
    output.push(b'\"');
    attr(&mut output, "name", &value.name);
    attr(&mut output, "sourceName", &value.source_name);
    if let Some(uid) = &value.uid {
        attr(&mut output, "xr10:uid", uid);
    }
    output.push(b'>');
    if !value.pivot_tables.is_empty() {
        output.extend_from_slice(b"<x15:pivotTables>");
        for pivot in &value.pivot_tables {
            output.extend_from_slice(b"<x15:pivotTable");
            attr(&mut output, "tabId", &pivot.tab_id.to_string());
            attr(&mut output, "name", &pivot.name);
            output.extend_from_slice(b"/>");
        }
        output.extend_from_slice(b"</x15:pivotTables>");
    }
    write_state(&mut output, &value.state)?;
    if let Some(payload) = &value.timeline_pivot_filter {
        write_timeline_pivot_filter(&mut output, payload)?;
    }
    if let Some(payload) = &value.extension_list {
        append_opaque_any_namespace(&mut output, payload, "extLst")?;
    }
    output.extend_from_slice(b"</x15:timelineCacheDefinition>");
    if output.len() > MAX_XML_BYTES {
        return Err(limit("serialized cache XML bytes"));
    }
    Ok(output)
}

/// Parse one worksheet-scoped Views part.
pub fn parse_timelines(xml: &[u8]) -> Result<Views> {
    let root = parse_document(xml)?;
    require(&root, X15, "timelines")?;
    no_attributes(&root, &[])?;
    whitespace(&root)?;
    if root.children.is_empty() {
        return Err(invalid("Views part must contain at least one timeline"));
    }
    if root.children.len() > MAX_TIMELINES {
        return Err(limit("timeline count"));
    }
    let mut timelines = Vec::with_capacity(root.children.len());
    for child in &root.children {
        timelines.push(parse_timeline(child)?);
    }
    let value = Views { timelines };
    validate_views_local(&value)?;
    Ok(value)
}

fn parse_timeline(node: &Node) -> Result<View> {
    require(node, X15, "timeline")?;
    whitespace(node)?;
    no_attributes(
        node,
        &[
            ("", "name"),
            (XR10, "uid"),
            ("", "cache"),
            ("", "caption"),
            ("", "showHeader"),
            ("", "showSelectionLabel"),
            ("", "showTimeLevel"),
            ("", "showHorizontalScrollbar"),
            ("", "level"),
            ("", "selectionLevel"),
            ("", "scrollPosition"),
            ("", "style"),
        ],
    )?;
    if node.children.len() > 1 {
        return Err(invalid("timeline permits at most one extLst"));
    }
    let extension_list = node
        .children
        .first()
        .map(|child| {
            if child.name != "extLst"
                || !(child.namespace == X15
                    || child.namespace == SML
                    || child.namespace == STRICT_SML)
            {
                return Err(invalid("timeline child must be extLst"));
            }
            opaque_from_node(child)
        })
        .transpose()?;
    Ok(View {
        name: required(node, "", "name")?.to_owned(),
        uid: optional(node, XR10, "uid").map(str::to_owned),
        cache: required(node, "", "cache")?.to_owned(),
        caption: optional(node, "", "caption").map(str::to_owned),
        show_header: optional_bool(node, "showHeader")?,
        show_selection_label: optional_bool(node, "showSelectionLabel")?,
        show_time_level: optional_bool(node, "showTimeLevel")?,
        show_horizontal_scrollbar: optional_bool(node, "showHorizontalScrollbar")?,
        level: Level::parse(required(node, "", "level")?, "level")?,
        selection_level: Level::parse(required(node, "", "selectionLevel")?, "selectionLevel")?,
        scroll_position: optional(node, "", "scrollPosition").map(str::to_owned),
        style: optional(node, "", "style").map(str::to_owned),
        extension_list,
    })
}

/// Deterministically serialize one worksheet-scoped Views part.
pub fn write_timelines(value: &Views) -> Result<Vec<u8>> {
    validate_views_local(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<x15:timelines xmlns:x15=\"");
    escape(&mut output, X15);
    output.extend_from_slice(b"\" xmlns:xr10=\"");
    escape(&mut output, XR10);
    output.extend_from_slice(b"\">");
    for timeline in &value.timelines {
        output.extend_from_slice(b"<x15:timeline");
        attr(&mut output, "name", &timeline.name);
        if let Some(uid) = &timeline.uid {
            attr(&mut output, "xr10:uid", uid);
        }
        attr(&mut output, "cache", &timeline.cache);
        if let Some(value) = &timeline.caption {
            attr(&mut output, "caption", value);
        }
        for (name, value) in [
            ("showHeader", timeline.show_header),
            ("showSelectionLabel", timeline.show_selection_label),
            ("showTimeLevel", timeline.show_time_level),
            (
                "showHorizontalScrollbar",
                timeline.show_horizontal_scrollbar,
            ),
        ] {
            if let Some(value) = value {
                bool_attr(&mut output, name, value);
            }
        }
        attr(&mut output, "level", timeline.level.number());
        attr(
            &mut output,
            "selectionLevel",
            timeline.selection_level.number(),
        );
        if let Some(value) = &timeline.scroll_position {
            attr(&mut output, "scrollPosition", value);
        }
        if let Some(value) = &timeline.style {
            attr(&mut output, "style", value);
        }
        if let Some(payload) = &timeline.extension_list {
            output.push(b'>');
            append_opaque_any_namespace(&mut output, payload, "extLst")?;
            output.extend_from_slice(b"</x15:timeline>");
        } else {
            output.extend_from_slice(b"/>");
        }
    }
    output.extend_from_slice(b"</x15:timelines>");
    if output.len() > MAX_XML_BYTES {
        return Err(limit("serialized timelines XML bytes"));
    }
    Ok(output)
}

/// Load and validate all workbook View Cache parts.
pub(super) fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("XML bytes"));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = NsReader::from_reader(xml);
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
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
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|value| value.to_string())
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
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside XML root"));
                }
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    root.ok_or_else(|| invalid("missing XML root"))
}

fn make_node(
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
            .any(|attribute: &Attribute| attribute.namespace == namespace && attribute.name == name)
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
    })
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

pub(super) fn opaque_from_node(node: &Node) -> Result<OpaqueXml> {
    let xml = serialize_node(node)?;
    if xml.len() > MAX_OPAQUE_BYTES {
        return Err(limit("opaque XML bytes"));
    }
    Ok(OpaqueXml { xml })
}

pub(super) fn serialize_node(node: &Node) -> Result<Vec<u8>> {
    let mut namespaces = BTreeMap::<String, String>::new();
    collect_namespaces(node, &mut namespaces);
    let mut prefixes = HashMap::new();
    let mut next = 0usize;
    for namespace in namespaces.keys() {
        let prefix = match namespace.as_str() {
            X15 => "x15".into(),
            SML | STRICT_SML => "x".into(),
            REL | STRICT_REL => "r".into(),
            XR10 => "xr10".into(),
            _ => {
                let value = format!("n{next}");
                next += 1;
                value
            },
        };
        prefixes.insert(namespace.clone(), prefix);
    }
    let mut output = Vec::new();
    write_node(node, &prefixes, true, &mut output);
    Ok(output)
}

fn collect_namespaces(node: &Node, output: &mut BTreeMap<String, String>) {
    if !node.namespace.is_empty() {
        output.insert(node.namespace.clone(), String::new());
    }
    for attr in &node.attributes {
        if !attr.namespace.is_empty() {
            output.insert(attr.namespace.clone(), String::new());
        }
    }
    for child in &node.children {
        collect_namespaces(child, output);
    }
}
fn write_node(node: &Node, prefixes: &HashMap<String, String>, root: bool, output: &mut Vec<u8>) {
    output.push(b'<');
    qname(output, &node.namespace, &node.name, prefixes);
    if root {
        let mut entries: Vec<_> = prefixes.iter().collect();
        entries.sort_by(|a, b| a.1.cmp(b.1));
        for (namespace, prefix) in entries {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b"=\"");
            escape(output, namespace);
            output.push(b'\"');
        }
    }
    for attr_value in &node.attributes {
        output.push(b' ');
        qname(output, &attr_value.namespace, &attr_value.name, prefixes);
        output.extend_from_slice(b"=\"");
        escape(output, &attr_value.value);
        output.push(b'\"');
    }
    if node.children.is_empty() && node.text.is_empty() {
        output.extend_from_slice(b"/>");
        return;
    }
    output.push(b'>');
    escape_text(output, &node.text);
    for child in &node.children {
        write_node(child, prefixes, false, output);
    }
    output.extend_from_slice(b"</");
    qname(output, &node.namespace, &node.name, prefixes);
    output.push(b'>');
}
fn qname(output: &mut Vec<u8>, namespace: &str, name: &str, prefixes: &HashMap<String, String>) {
    if !namespace.is_empty() {
        output.extend_from_slice(prefixes[namespace].as_bytes());
        output.push(b':');
    }
    output.extend_from_slice(name.as_bytes());
}

fn append_opaque_any_namespace(
    output: &mut Vec<u8>,
    payload: &OpaqueXml,
    name: &str,
) -> Result<()> {
    validate_opaque_kind(payload, &[X15, SML, STRICT_SML], name)?;
    output.extend_from_slice(&payload.xml);
    Ok(())
}
pub(super) fn validate_opaque_kind(
    payload: &OpaqueXml,
    namespaces: &[&str],
    name: &str,
) -> Result<()> {
    if payload.xml.len() > MAX_OPAQUE_BYTES {
        return Err(limit("opaque XML bytes"));
    }
    let root = parse_document(&payload.xml)?;
    if root.name != name || !namespaces.contains(&root.namespace.as_str()) {
        return Err(invalid(format!(
            "opaque XML must be a {name} element in its normative namespace"
        )));
    }
    Ok(())
}

pub(super) fn require(node: &Node, namespace: &str, name: &str) -> Result<()> {
    if node.namespace == namespace && node.name == name {
        Ok(())
    } else {
        Err(invalid(format!(
            "expected {{{namespace}}}{name}, got {{{}}}{}",
            node.namespace, node.name
        )))
    }
}
pub(super) fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}
pub(super) fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> {
    optional(node, namespace, name)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid(format!("{} is missing attribute '{name}'", node.name)))
}
pub(super) fn optional_bool(node: &Node, name: &str) -> Result<Option<bool>> {
    optional(node, "", name)
        .map(|value| parse_bool(value, name))
        .transpose()
}
fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid boolean '{value}' for {name}"))),
    }
}
pub(super) fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    if let Some(attribute) = node.attributes.iter().find(|attribute| {
        !allowed.contains(&(attribute.namespace.as_str(), attribute.name.as_str()))
    }) {
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
pub(super) fn empty(node: &Node) -> Result<()> {
    whitespace(node)?;
    if node.children.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{} must be empty", node.name)))
    }
}
fn add_strings(total: &mut usize, size: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
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
fn bool_attr(output: &mut Vec<u8>, name: &str, value: bool) {
    attr(output, name, if value { "1" } else { "0" });
}
pub(super) fn attr(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape(output, value);
    output.push(b'\"');
}
pub(super) fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
pub(super) fn escape_text(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
