#![allow(
    clippy::map_err_ignore,
    clippy::wildcard_enum_match_arm,
    reason = "legacy module confines normalization into the module's stable typed public error, an intentional opaque or future-variant fallback to this codec boundary"
)]

//! Bounded XML codec for XLSB timeline cache and view parts.

use super::model::{Cache, Filter, FilterType, Level, PivotTable, Range, State, View, Views};
use super::validation::{cache as validate_cache, views as validate_views};
use crate::package::error::{Error, Result};
use litchi_core::xml::escape_xml;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::BTreeMap;

pub(crate) const X15: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
pub(crate) const XR10: &str = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";
pub(crate) const CACHE_CONTENT_TYPE: &str = "application/vnd.ms-excel.TimelineCache+xml";
pub(crate) const CACHE_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2010/relationships/TimelineCache";
pub(crate) const VIEWS_CONTENT_TYPE: &str = "application/vnd.ms-excel.Timeline+xml";
pub(crate) const VIEWS_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2010/relationships/Timeline";

const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_NODES: usize = 250_000;
const MAX_DEPTH: usize = 128;

#[derive(Debug, Clone)]
struct Node {
    namespace: String,
    name: String,
    attributes: BTreeMap<String, String>,
    children: Vec<Node>,
    text: String,
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn element_name(bytes: &[u8]) -> Result<(String, String)> {
    let text = std::str::from_utf8(bytes)
        .map_err(|error| invalid(format!("invalid XML name: {error}")))?;
    let (prefix, local) = text
        .split_once(':')
        .map_or(("", text), |(prefix, local)| (prefix, local));
    let namespace = match prefix {
        "" | "x15" => X15,
        "xr10" => XR10,
        _ => {
            return Err(invalid(format!(
                "unsupported timeline XML prefix '{prefix}'"
            )));
        },
    };
    Ok((namespace.to_string(), local.to_string()))
}

fn attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<BTreeMap<String, String>> {
    let mut result = BTreeMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid XML attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let key = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("invalid XML attribute name: {error}")))?
            .to_string();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid XML attribute value: {error}")))?
            .into_owned();
        if result.insert(key.clone(), value).is_some() {
            return Err(invalid(format!("duplicate XML attribute '{key}'")));
        }
    }
    Ok(result)
}

fn resolve_namespace(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<String> {
    match reader.resolver().resolve_element(element.name()).0 {
        ResolveResult::Bound(Namespace(namespace)) => String::from_utf8(namespace.to_vec())
            .map_err(|error| invalid(format!("invalid XML namespace: {error}"))),
        _ => Err(invalid("timeline XML element has no bound namespace")),
    }
}

fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("timeline XML exceeds the bounded byte limit"));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut closed = false;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid timeline XML: {error}")))?
            .into_owned();
        let empty_event = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(invalid("timeline XML structure limit exceeded"));
                }
                let (namespace, name) = element_name(element.name().as_ref())?;
                let resolved = resolve_namespace(&reader, &element)?;
                if resolved != namespace {
                    return Err(invalid(format!(
                        "timeline XML prefix '{name}' is bound to an unexpected namespace"
                    )));
                }
                let node = Node {
                    namespace,
                    name,
                    attributes: attributes(&reader, &element)?,
                    children: Vec::new(),
                    text: String::new(),
                };
                if empty_event {
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(node);
                    } else if root.replace(node).is_some() {
                        return Err(invalid("timeline XML contains multiple roots"));
                    } else {
                        closed = true;
                    }
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected timeline XML closing element"))?;
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(node);
                } else if root.replace(node).is_some() {
                    return Err(invalid("timeline XML contains multiple roots"));
                } else {
                    closed = true;
                }
            },
            Event::Text(text) => {
                let value = text
                    .decode()
                    .map_err(|error| invalid(format!("invalid timeline XML text: {error}")))?;
                let value = quick_xml::escape::unescape(&value)
                    .map_err(|error| invalid(format!("invalid timeline XML entity: {error}")))?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else if !value.trim().is_empty() {
                    return Err(invalid("timeline XML contains text outside its root"));
                }
            },
            Event::GeneralRef(reference) => {
                let value = reference
                    .resolve_char_ref()
                    .map_err(|error| invalid(format!("invalid XML reference: {error}")))?
                    .map(|value| value.to_string())
                    .ok_or_else(|| invalid("custom XML entities are not supported"))?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("timeline XML entity outside root"));
                }
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::CData(_) | Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("timeline XML contains unsupported event"));
            },
            Event::Eof => break,
        }
    }
    if !stack.is_empty() || !closed {
        return Err(invalid("timeline XML is unterminated or has no root"));
    }
    root.ok_or_else(|| invalid("timeline XML has no root"))
}

fn allowed(node: &Node, names: &[&str]) -> Result<()> {
    if node
        .attributes
        .keys()
        .any(|key| !names.contains(&key.as_str()))
    {
        return Err(invalid(format!(
            "unexpected attribute on '{}': {:?}",
            node.name, node.attributes
        )));
    }
    Ok(())
}

fn attr<'a>(node: &'a Node, name: &str, required: bool) -> Result<Option<&'a str>> {
    let value = node.attributes.get(name).map(String::as_str);
    if required && value.is_none() {
        return Err(invalid(format!(
            "'{}' requires attribute '{name}'",
            node.name
        )));
    }
    Ok(value)
}

fn required_attr<'a>(node: &'a Node, name: &str) -> Result<&'a str> {
    attr(node, name, true)?
        .ok_or_else(|| invalid(format!("'{}' requires attribute '{name}'", node.name)))
}

fn empty(node: &Node) -> Result<()> {
    if !node.children.is_empty() || !node.text.trim().is_empty() {
        return Err(invalid(format!(
            "timeline XML element '{}' must be empty",
            node.name
        )));
    }
    Ok(())
}

fn parse_bool(value: Option<&str>, default: bool, field: &str) -> Result<bool> {
    match value {
        None => Ok(default),
        Some("1" | "true") => Ok(true),
        Some("0" | "false") => Ok(false),
        Some(value) => Err(invalid(format!("invalid {field} value '{value}'"))),
    }
}

fn parse_u32(node: &Node, name: &str) -> Result<u32> {
    required_attr(node, name)?
        .parse()
        .map_err(|_| invalid(format!("invalid {name} value")))
}

fn parse_range(node: &Node) -> Result<Range> {
    if node.namespace != X15 || !matches!(node.name.as_str(), "selection" | "bounds") {
        return Err(invalid("invalid timeline range element"));
    }
    allowed(node, &["startDate", "endDate"])?;
    empty(node)?;
    Range::new(
        required_attr(node, "startDate")?,
        required_attr(node, "endDate")?,
    )
}

fn parse_state(node: &Node) -> Result<State> {
    if node.namespace != X15 || node.name != "state" {
        return Err(invalid("timeline cache requires x15:state"));
    }
    allowed(
        node,
        &[
            "singleRangeFilterState",
            "minimalRefreshVersion",
            "lastRefreshVersion",
            "pivotCacheId",
            "filterType",
        ],
    )?;
    let mut selection = None;
    let mut bounds = None;
    for child in &node.children {
        if child.name == "selection" && selection.is_none() {
            selection = Some(parse_range(child)?);
        } else if child.name == "bounds" && bounds.is_none() {
            bounds = Some(parse_range(child)?);
        } else {
            return Err(invalid(
                "timeline state child order or cardinality is invalid",
            ));
        }
    }
    let value = State {
        selection,
        bounds: bounds.ok_or_else(|| invalid("timeline state requires bounds"))?,
        single_range_filter_state: parse_bool(
            attr(node, "singleRangeFilterState", false)?,
            true,
            "singleRangeFilterState",
        )?,
        minimal_refresh_version: parse_u32(node, "minimalRefreshVersion")?,
        last_refresh_version: parse_u32(node, "lastRefreshVersion")?,
        pivot_cache_id: parse_u32(node, "pivotCacheId")?,
        filter_type: FilterType::parse(required_attr(node, "filterType")?)?,
    };
    super::model::validate_state(&value)?;
    Ok(value)
}

fn parse_filter(node: &Node) -> Result<Filter> {
    if node.namespace != X15 || node.name != "timelinePivotFilter" {
        return Err(invalid("invalid timelinePivotFilter"));
    }
    allowed(node, &["useWholeDay", "fld", "id", "name", "description"])?;
    empty(node)?;
    Ok(Filter {
        use_whole_day: parse_bool(attr(node, "useWholeDay", false)?, false, "useWholeDay")?,
        field: parse_u32(node, "fld")?,
        id: parse_u32(node, "id")?,
        name: attr(node, "name", false)?.map(str::to_owned),
        description: attr(node, "description", false)?.map(str::to_owned),
    })
}

/// Parse one timeline cache definition XML part.
pub fn parse_cache(xml: &[u8]) -> Result<Cache> {
    let root = parse_document(xml)?;
    if root.namespace != X15 || root.name != "timelineCacheDefinition" {
        return Err(invalid(
            "timeline cache root must be x15:timelineCacheDefinition",
        ));
    }
    allowed(&root, &["name", "sourceName", "xr10:uid"])?;
    let mut pivot_tables = Vec::new();
    let mut state = None;
    let mut filter = None;
    let mut stage = 0u8;
    for child in &root.children {
        match child.name.as_str() {
            "pivotTables" if stage == 0 => {
                if child.namespace != X15
                    || !child.attributes.is_empty()
                    || !child.text.trim().is_empty()
                {
                    return Err(invalid("invalid timeline pivotTables"));
                }
                for pivot in &child.children {
                    if pivot.namespace != X15 || pivot.name != "pivotTable" {
                        return Err(invalid("invalid timeline pivotTable child"));
                    }
                    allowed(pivot, &["tabId", "name"])?;
                    empty(pivot)?;
                    pivot_tables.push(PivotTable {
                        tab_id: parse_u32(pivot, "tabId")?,
                        name: required_attr(pivot, "name")?.to_string(),
                    });
                }
                stage = 1;
            },
            "state" if stage <= 1 && state.is_none() => {
                state = Some(parse_state(child)?);
                stage = 2;
            },
            "timelinePivotFilter" if stage == 2 && filter.is_none() => {
                filter = Some(parse_filter(child)?);
                stage = 3;
            },
            _ => return Err(invalid("invalid or out-of-order timeline cache child")),
        }
    }
    let cache = Cache {
        name: required_attr(&root, "name")?.to_string(),
        source_name: required_attr(&root, "sourceName")?.to_string(),
        uid: attr(&root, "xr10:uid", false)?.map(str::to_owned),
        pivot_tables,
        state: state.ok_or_else(|| invalid("timeline cache requires exactly one state"))?,
        filter,
    };
    validate_cache(&cache)?;
    Ok(cache)
}

fn parse_view(node: &Node) -> Result<View> {
    if node.namespace != X15 || node.name != "timeline" {
        return Err(invalid("timeline child must be x15:timeline"));
    }
    allowed(
        node,
        &[
            "name",
            "cache",
            "level",
            "selectionLevel",
            "caption",
            "showHeader",
            "showSelectionLabel",
            "showTimeLevel",
            "showHorizontalScrollbar",
            "scrollPosition",
            "style",
            "xr10:uid",
        ],
    )?;
    empty(node)?;
    Ok(View {
        name: required_attr(node, "name")?.to_string(),
        cache: required_attr(node, "cache")?.to_string(),
        level: Level::parse(required_attr(node, "level")?)?,
        selection_level: Level::parse(required_attr(node, "selectionLevel")?)?,
        caption: attr(node, "caption", false)?.map(str::to_owned),
        show_header: parse_bool(attr(node, "showHeader", false)?, true, "showHeader")?,
        show_selection_label: parse_bool(
            attr(node, "showSelectionLabel", false)?,
            true,
            "showSelectionLabel",
        )?,
        show_time_level: parse_bool(attr(node, "showTimeLevel", false)?, true, "showTimeLevel")?,
        show_horizontal_scrollbar: parse_bool(
            attr(node, "showHorizontalScrollbar", false)?,
            true,
            "showHorizontalScrollbar",
        )?,
        scroll_position: attr(node, "scrollPosition", false)?.map(str::to_owned),
        style: attr(node, "style", false)?.map(str::to_owned),
        uid: attr(node, "xr10:uid", false)?.map(str::to_owned),
    })
}

/// Parse one worksheet timelines XML part.
pub fn parse_views(xml: &[u8]) -> Result<Views> {
    let root = parse_document(xml)?;
    if root.namespace != X15 || root.name != "timelines" || !root.attributes.is_empty() {
        return Err(invalid(
            "timeline views root must be x15:timelines without attributes",
        ));
    }
    let mut value = Views::new();
    for child in &root.children {
        value.items.push(parse_view(child)?);
    }
    validate_views(&value)?;
    Ok(value)
}

fn attr_write(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&escape_xml(value));
    output.push('"');
}

fn bool_write(output: &mut String, name: &str, value: bool) {
    attr_write(output, name, if value { "1" } else { "0" });
}

fn range_write(output: &mut String, name: &str, value: &Range) {
    output.push_str("<x15:");
    output.push_str(name);
    attr_write(output, "startDate", &value.start_date);
    attr_write(output, "endDate", &value.end_date);
    output.push_str("/>");
}

/// Encode one timeline cache definition.
pub fn write_cache(value: &Cache) -> Result<Vec<u8>> {
    validate_cache(value)?;
    let mut output = String::from("<x15:timelineCacheDefinition xmlns:x15=\"");
    output.push_str(X15);
    output.push_str("\" xmlns:xr10=\"");
    output.push_str(XR10);
    output.push('"');
    attr_write(&mut output, "name", &value.name);
    attr_write(&mut output, "sourceName", &value.source_name);
    if let Some(uid) = &value.uid {
        attr_write(&mut output, "xr10:uid", uid);
    }
    output.push('>');
    if !value.pivot_tables.is_empty() {
        output.push_str("<x15:pivotTables>");
        for pivot in &value.pivot_tables {
            output.push_str("<x15:pivotTable");
            attr_write(&mut output, "tabId", &pivot.tab_id.to_string());
            attr_write(&mut output, "name", &pivot.name);
            output.push_str("/>");
        }
        output.push_str("</x15:pivotTables>");
    }
    let state = &value.state;
    output.push_str("<x15:state");
    bool_write(
        &mut output,
        "singleRangeFilterState",
        state.single_range_filter_state,
    );
    attr_write(
        &mut output,
        "minimalRefreshVersion",
        &state.minimal_refresh_version.to_string(),
    );
    attr_write(
        &mut output,
        "lastRefreshVersion",
        &state.last_refresh_version.to_string(),
    );
    attr_write(
        &mut output,
        "pivotCacheId",
        &state.pivot_cache_id.to_string(),
    );
    attr_write(&mut output, "filterType", state.filter_type.as_str());
    output.push('>');
    if let Some(selection) = &state.selection {
        range_write(&mut output, "selection", selection);
    }
    range_write(&mut output, "bounds", &state.bounds);
    output.push_str("</x15:state>");
    if let Some(filter) = &value.filter {
        output.push_str("<x15:timelinePivotFilter");
        bool_write(&mut output, "useWholeDay", filter.use_whole_day);
        attr_write(&mut output, "fld", &filter.field.to_string());
        attr_write(&mut output, "id", &filter.id.to_string());
        if let Some(name) = &filter.name {
            attr_write(&mut output, "name", name);
        }
        if let Some(description) = &filter.description {
            attr_write(&mut output, "description", description);
        }
        output.push_str("/>");
    }
    output.push_str("</x15:timelineCacheDefinition>");
    if output.len() > MAX_XML_BYTES {
        return Err(invalid("timeline cache XML exceeds the bounded byte limit"));
    }
    Ok(output.into_bytes())
}

/// Encode one worksheet timelines XML part.
pub fn write_views(value: &Views) -> Result<Vec<u8>> {
    validate_views(value)?;
    let mut output = String::from("<x15:timelines xmlns:x15=\"");
    output.push_str(X15);
    output.push_str("\" xmlns:xr10=\"");
    output.push_str(XR10);
    output.push_str("\">");
    for view in &value.items {
        output.push_str("<x15:timeline");
        attr_write(&mut output, "name", &view.name);
        attr_write(&mut output, "cache", &view.cache);
        attr_write(&mut output, "level", &view.level.wire().to_string());
        attr_write(
            &mut output,
            "selectionLevel",
            &view.selection_level.wire().to_string(),
        );
        if let Some(caption) = &view.caption {
            attr_write(&mut output, "caption", caption);
        }
        bool_write(&mut output, "showHeader", view.show_header);
        bool_write(&mut output, "showSelectionLabel", view.show_selection_label);
        bool_write(&mut output, "showTimeLevel", view.show_time_level);
        bool_write(
            &mut output,
            "showHorizontalScrollbar",
            view.show_horizontal_scrollbar,
        );
        if let Some(scroll) = &view.scroll_position {
            attr_write(&mut output, "scrollPosition", scroll);
        }
        if let Some(style) = &view.style {
            attr_write(&mut output, "style", style);
        }
        if let Some(uid) = &view.uid {
            attr_write(&mut output, "xr10:uid", uid);
        }
        output.push_str("/>");
    }
    output.push_str("</x15:timelines>");
    if output.len() > MAX_XML_BYTES {
        return Err(invalid("timeline views XML exceeds the bounded byte limit"));
    }
    Ok(output.into_bytes())
}
