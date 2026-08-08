//! Bounded namespace-aware XML codec for ODF handout masters.

use super::Master;
use super::validation::{
    MAX_CHILDREN, MAX_FRAGMENT_BYTES, MAX_TOTAL_CHILD_BYTES, MAX_XML_BYTES, invalid,
};
use litchi_core::Result;
use litchi_odf_common::style::master::{Child, ChildKind};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const DR3D: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
const SCRIPT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const SVG: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const MAX_DEPTH: usize = 256;
const MAX_EVENTS: usize = 1_000_000;

#[derive(Clone, Debug)]
struct Name {
    namespace: Option<Vec<u8>>,
    local: Vec<u8>,
}

#[derive(Clone, Debug)]
struct Location {
    start: usize,
    open_end: usize,
    content_end: usize,
    end: usize,
    qname: String,
    empty: bool,
}

#[derive(Clone, Debug)]
struct Parsed {
    master: Option<Master>,
    master_location: Option<Location>,
    container: Option<Location>,
}

#[derive(Clone, Debug)]
struct Active {
    page_layout_name: String,
    presentation_page_layout_name: Option<String>,
    drawing_style_name: Option<String>,
    header_name: Option<String>,
    footer_name: Option<String>,
    date_time_name: Option<String>,
    start: usize,
    depth: usize,
    children: Vec<Child>,
    child: Option<(usize, usize)>,
    child_bytes: usize,
}

#[derive(Clone, Debug)]
struct Attribute {
    qname: String,
    namespace: Option<Vec<u8>>,
    local: String,
    value: String,
}

#[derive(Clone, Debug)]
struct FragmentParts {
    qname: String,
    open: String,
    inner: String,
}

fn bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn expanded(reader: &NsReader<&[u8]>, name: quick_xml::name::QName<'_>) -> Result<Name> {
    let (namespace, local) = reader.resolver().resolve_element(name);
    let namespace = match namespace {
        ResolveResult::Bound(Namespace(value)) => Some(value.to_vec()),
        ResolveResult::Unbound => None,
        ResolveResult::Unknown(prefix) => {
            return Err(invalid(format!(
                "unbound handout-master element prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            )));
        },
    };
    Ok(Name {
        namespace,
        local: local.as_ref().to_vec(),
    })
}

fn is_name(name: &Name, namespace: &[u8], local: &[u8]) -> bool {
    name.namespace.as_deref() == Some(namespace) && name.local == local
}

fn is_shape(name: &Name) -> bool {
    if is_name(name, DR3D, b"scene") {
        return true;
    }
    name.namespace.as_deref() == Some(DRAW)
        && matches!(
            name.local.as_slice(),
            b"a" | b"caption"
                | b"circle"
                | b"connector"
                | b"control"
                | b"custom-shape"
                | b"ellipse"
                | b"frame"
                | b"g"
                | b"line"
                | b"measure"
                | b"page-thumbnail"
                | b"path"
                | b"polygon"
                | b"polyline"
                | b"rect"
                | b"regular-polygon"
        )
}

fn parse_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    let mut found = None;
    for raw in element.attributes().with_checks(true) {
        let raw =
            raw.map_err(|error| invalid(format!("invalid handout-master attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(raw.key);
        if bound(&namespace, expected_namespace) && local.as_ref() == local_name {
            if found.is_some() {
                return Err(invalid(format!(
                    "duplicate handout-master attribute '{}',",
                    String::from_utf8_lossy(local_name)
                )));
            }
            let value = raw
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    invalid(format!("invalid handout-master attribute value: {error}"))
                })?
                .into_owned();
            found = Some(value);
        }
    }
    Ok(found)
}

fn root_fields(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<ActiveFields> {
    let page_layout_name = parse_attribute(reader, element, STYLE, b"page-layout-name")?
        .ok_or_else(|| invalid("style:handout-master is missing style:page-layout-name"))?;
    let mut expanded = HashSet::new();
    for raw in element.attributes().with_checks(true) {
        let raw =
            raw.map_err(|error| invalid(format!("invalid handout-master attribute: {error}")))?;
        if raw.key.as_ref() == b"xmlns" || raw.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(raw.key);
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(value)) => Some(value.to_vec()),
            ResolveResult::Unbound => None,
            ResolveResult::Unknown(prefix) => {
                return Err(invalid(format!(
                    "unbound handout-master attribute prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        };
        let key = (namespace, local.as_ref().to_vec());
        if !expanded.insert(key) {
            return Err(invalid("duplicate expanded handout-master attribute"));
        }
    }
    Ok(ActiveFields {
        page_layout_name,
        presentation_page_layout_name: parse_attribute(
            reader,
            element,
            PRESENTATION,
            b"presentation-page-layout-name",
        )?,
        drawing_style_name: parse_attribute(reader, element, DRAW, b"style-name")?,
        header_name: parse_attribute(reader, element, PRESENTATION, b"use-header-name")?,
        footer_name: parse_attribute(reader, element, PRESENTATION, b"use-footer-name")?,
        date_time_name: parse_attribute(reader, element, PRESENTATION, b"use-date-time-name")?,
    })
}

#[derive(Clone, Debug)]
struct ActiveFields {
    page_layout_name: String,
    presentation_page_layout_name: Option<String>,
    drawing_style_name: Option<String>,
    header_name: Option<String>,
    footer_name: Option<String>,
    date_time_name: Option<String>,
}

/// Read the optional singleton handout master from a styles XML part.
pub(crate) fn read(xml: &str) -> Result<Option<Master>> {
    parse_document(xml).map(|parsed| parsed.master)
}

pub(crate) fn parse_fragment(xml: &str) -> Result<Master> {
    if xml.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("handout-master fragment exceeds 16 MiB"));
    }
    let wrapped = format!(
        "<office:document-styles xmlns:office=\"{}\"><office:master-styles>{xml}</office:master-styles></office:document-styles>",
        String::from_utf8_lossy(OFFICE)
    );
    let parsed = parse_document(&wrapped)?;
    let master = parsed
        .master
        .ok_or_else(|| invalid("fragment is missing style:handout-master"))?;
    if master.source != xml {
        return Err(invalid(
            "fragment must contain exactly one style:handout-master element",
        ));
    }
    Ok(master)
}

pub(crate) fn validate_shape_fragment(xml: &str) -> Result<()> {
    if xml.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("handout drawing child exceeds 16 MiB"));
    }
    let wrapped = format!(
        "<office:document-styles xmlns:office=\"{}\"><office:master-styles><style:handout-master xmlns:style=\"{}\" xmlns:draw=\"{}\" xmlns:presentation=\"{}\" xmlns:dr3d=\"{}\" xmlns:svg=\"{}\" xmlns:text=\"{}\" xmlns:xlink=\"{}\" style:page-layout-name=\"layout\">{xml}</style:handout-master></office:master-styles></office:document-styles>",
        String::from_utf8_lossy(OFFICE),
        String::from_utf8_lossy(STYLE),
        String::from_utf8_lossy(DRAW),
        String::from_utf8_lossy(PRESENTATION),
        String::from_utf8_lossy(DR3D),
        String::from_utf8_lossy(SVG),
        String::from_utf8_lossy(TEXT),
        String::from_utf8_lossy(XLINK)
    );
    let parsed = parse_document(&wrapped)?;
    let master = parsed
        .master
        .ok_or_else(|| invalid("drawing fragment did not parse"))?;
    if master.children.len() != 1 || master.children[0].xml != xml {
        return Err(invalid(
            "drawing child must be exactly one supported direct handout shape",
        ));
    }
    Ok(())
}

pub(crate) fn write(master: &Master) -> Result<String> {
    super::validation::validate(master)?;
    if !master.source.is_empty()
        && let Ok(original) = parse_fragment(&master.source)
        && original == *master
    {
        return Ok(master.source.clone());
    }

    let (open, inner, qname) = if master.source.is_empty() {
        (
            canonical_open(master),
            master
                .children
                .iter()
                .map(|child| child.xml.as_str())
                .collect::<String>(),
            "style:handout-master".to_string(),
        )
    } else {
        let parts = fragment_parts(&master.source)?;
        let open = render_open(&parts.open, master)?;
        let inner = if parse_fragment(&master.source)
            .map(|original| original.children == master.children)
            .unwrap_or(false)
        {
            parts.inner
        } else {
            master
                .children
                .iter()
                .map(|child| child.xml.as_str())
                .collect::<String>()
        };
        (open, inner, parts.qname)
    };
    let mut output = open;
    if inner.is_empty() {
        if output.ends_with('>') {
            output.pop();
            output.push_str("/>");
        } else {
            return Err(invalid("handout-master open tag is malformed"));
        }
    } else {
        if !output.ends_with('>') {
            output.push('>');
        }
        output.push_str(&inner);
        output.push_str("</");
        output.push_str(&qname);
        output.push('>');
    }
    let parsed = parse_fragment(&output)?;
    if parsed != *master {
        return Err(invalid("serialized handout master did not roundtrip"));
    }
    Ok(output)
}

pub(crate) fn replace_in_styles(styles: &str, fragment: &str) -> Result<String> {
    let fragment = parse_fragment(fragment)?;
    let fragment = fragment.source;
    let (bom, body) = split_bom(styles);
    let parsed = parse_document(body)?;
    let body = if let Some(location) = parsed.master_location {
        replace_range(body, location.start, location.end, &fragment)?
    } else {
        let container = parsed
            .container
            .ok_or_else(|| invalid("styles XML is missing office:master-styles"))?;
        insert_into_container(body, &container, &fragment)?
    };
    Ok(format!("{bom}{body}"))
}

pub(crate) fn remove_from_styles(styles: &str) -> Result<String> {
    let (bom, body) = split_bom(styles);
    let parsed = parse_document(body)?;
    let Some(location) = parsed.master_location else {
        return Ok(styles.to_string());
    };
    let body = replace_range(body, location.start, location.end, "")?;
    Ok(format!("{bom}{body}"))
}

fn parse_document(xml: &str) -> Result<Parsed> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("handout-master XML exceeds 64 MiB"));
    }
    let xml = xml.strip_prefix('\u{feff}').unwrap_or(xml);
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Name>::new();
    let mut active = None;
    let mut master = None;
    let mut master_location = None;
    let mut container = None;
    let mut container_open = Vec::<(usize, usize, String, bool)>::new();
    let mut events = 0usize;

    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(invalid("handout-master XML has too many events"));
        }
        let event_start = reader.buffer_position() as usize;
        let (_namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("handout-master XML parsing error: {error}")))?;
        let event_end = reader.buffer_position() as usize;
        let event = event.into_owned();
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("handout-master XML exceeds 256 levels"));
                }
                let name = expanded(&reader, element.name())?;
                reject_active(&name)?;
                let depth = stack.len() + 1;
                if is_name(&name, OFFICE, b"master-styles") {
                    if container.is_some()
                        || !stack
                            .iter()
                            .any(|item| is_name(item, OFFICE, b"document-styles"))
                    {
                        return Err(invalid(
                            "office:master-styles is not a valid styles container",
                        ));
                    }
                    container_open.push((event_start, event_end, element_name(&element)?, false));
                }
                if is_name(&name, STYLE, b"handout-master") {
                    let parent = stack.last().ok_or_else(|| {
                        invalid("style:handout-master must be inside office:master-styles")
                    })?;
                    if !is_name(parent, OFFICE, b"master-styles")
                        || active.is_some()
                        || master.is_some()
                    {
                        return Err(invalid(
                            "styles XML must contain at most one direct style:handout-master",
                        ));
                    }
                    let fields = root_fields(&reader, &element)?;
                    active = Some(Active {
                        page_layout_name: fields.page_layout_name,
                        presentation_page_layout_name: fields.presentation_page_layout_name,
                        drawing_style_name: fields.drawing_style_name,
                        header_name: fields.header_name,
                        footer_name: fields.footer_name,
                        date_time_name: fields.date_time_name,
                        start: event_start,
                        depth,
                        children: Vec::new(),
                        child: None,
                        child_bytes: 0,
                    });
                } else if let Some(active) = active.as_mut()
                    && depth == active.depth + 1
                {
                    if active.child.is_some() {
                        return Err(invalid("handout drawing children overlap"));
                    }
                    if !is_shape(&name) {
                        return Err(invalid("unsupported direct style:handout-master child"));
                    }
                    active.child = Some((event_start, depth));
                }
                stack.push(name);
            },
            Event::Empty(element) => {
                let name = expanded(&reader, element.name())?;
                reject_active(&name)?;
                let depth = stack.len() + 1;
                if is_name(&name, OFFICE, b"master-styles") {
                    if container.is_some()
                        || !stack
                            .iter()
                            .any(|item| is_name(item, OFFICE, b"document-styles"))
                    {
                        return Err(invalid(
                            "office:master-styles is not a valid styles container",
                        ));
                    }
                    container = Some(Location {
                        start: event_start,
                        open_end: event_end,
                        content_end: event_start,
                        end: event_end,
                        qname: element_name(&element)?,
                        empty: true,
                    });
                } else if is_name(&name, STYLE, b"handout-master") {
                    let parent = stack.last().ok_or_else(|| {
                        invalid("style:handout-master must be inside office:master-styles")
                    })?;
                    if !is_name(parent, OFFICE, b"master-styles")
                        || active.is_some()
                        || master.is_some()
                    {
                        return Err(invalid(
                            "styles XML must contain at most one direct style:handout-master",
                        ));
                    }
                    let fields = root_fields(&reader, &element)?;
                    let value = Master {
                        page_layout_name: fields.page_layout_name,
                        presentation_page_layout_name: fields.presentation_page_layout_name,
                        drawing_style_name: fields.drawing_style_name,
                        header_name: fields.header_name,
                        footer_name: fields.footer_name,
                        date_time_name: fields.date_time_name,
                        children: Vec::new(),
                        source: xml[event_start..event_end].to_string(),
                    };
                    super::validation::validate_fields(&value)?;
                    master_location = Some(Location {
                        start: event_start,
                        open_end: event_end,
                        content_end: event_start,
                        end: event_end,
                        qname: element_name(&element)?,
                        empty: true,
                    });
                    master = Some(value);
                } else if let Some(active) = active.as_mut()
                    && depth == active.depth + 1
                {
                    if !is_shape(&name) {
                        return Err(invalid("unsupported direct style:handout-master child"));
                    }
                    push_child(active, xml, event_start, event_end)?;
                }
            },
            Event::End(element) => {
                let depth = stack.len();
                if let Some(active) = active.as_mut()
                    && active
                        .child
                        .as_ref()
                        .is_some_and(|(_, child_depth)| *child_depth == depth)
                {
                    let (start, _) = active.child.take().ok_or_else(|| {
                        invalid("handout-master child state disappeared during close")
                    })?;
                    push_child(active, xml, start, event_end)?;
                }
                if let Some(active_value) = active.as_ref()
                    && active_value.depth == depth
                {
                    let active_value = active
                        .take()
                        .ok_or_else(|| invalid("handout-master state disappeared during close"))?;
                    let source = xml[active_value.start..event_end].to_string();
                    let value = Master {
                        page_layout_name: active_value.page_layout_name,
                        presentation_page_layout_name: active_value.presentation_page_layout_name,
                        drawing_style_name: active_value.drawing_style_name,
                        header_name: active_value.header_name,
                        footer_name: active_value.footer_name,
                        date_time_name: active_value.date_time_name,
                        children: active_value.children,
                        source,
                    };
                    super::validation::validate_fields(&value)?;
                    master_location = Some(Location {
                        start: active_value.start,
                        open_end: find_open_end(xml, active_value.start, event_start)?,
                        content_end: event_start,
                        end: event_end,
                        qname: end_name(&element)?,
                        empty: false,
                    });
                    master = Some(value);
                }
                let current = stack
                    .pop()
                    .ok_or_else(|| invalid("handout XML depth underflow"))?;
                if is_name(&current, OFFICE, b"master-styles") {
                    let (start, open_end, qname, empty) = container_open
                        .pop()
                        .ok_or_else(|| invalid("master-styles container stack underflow"))?;
                    container = Some(Location {
                        start,
                        open_end,
                        content_end: event_start,
                        end: event_end,
                        qname,
                        empty,
                    });
                }
            },
            Event::Text(value) => {
                let bytes = value.as_ref();
                if stack.is_empty() && !bytes.iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("non-whitespace text is outside the XML document"));
                }
                if let Some(active) = &active
                    && stack.len() == active.depth
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid(
                        "non-whitespace text is not allowed directly in handout-master",
                    ));
                }
            },
            Event::CData(value) => {
                let bytes = value.as_ref();
                if stack.is_empty() && !bytes.iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("non-whitespace CDATA is outside the XML document"));
                }
                if let Some(active) = &active
                    && stack.len() == active.depth
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("CDATA is not allowed directly in handout-master"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are not allowed"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if active.is_some() || !stack.is_empty() || !container_open.is_empty() {
        return Err(invalid("truncated handout-master XML"));
    }
    Ok(Parsed {
        master,
        master_location,
        container,
    })
}

fn push_child(active: &mut Active, xml: &str, start: usize, end: usize) -> Result<()> {
    let size = end
        .checked_sub(start)
        .ok_or_else(|| invalid("invalid handout child span"))?;
    if size > MAX_FRAGMENT_BYTES {
        return Err(invalid("handout drawing child exceeds 16 MiB"));
    }
    active.child_bytes = active
        .child_bytes
        .checked_add(size)
        .ok_or_else(|| invalid("handout child size overflows"))?;
    if active.child_bytes > MAX_TOTAL_CHILD_BYTES {
        return Err(invalid("handout drawing children exceed 32 MiB"));
    }
    if active.children.len() >= MAX_CHILDREN {
        return Err(invalid("handout master has too many drawing children"));
    }
    active.children.push(Child::new(
        ChildKind::Shape,
        xml.get(start..end)
            .ok_or_else(|| invalid("handout child span is not UTF-8 aligned"))?,
    ));
    Ok(())
}

fn reject_active(name: &Name) -> Result<()> {
    if name.namespace.as_deref() == Some(SCRIPT)
        || (name.namespace.as_deref() == Some(OFFICE)
            && matches!(name.local.as_slice(), b"scripts" | b"event-listeners"))
    {
        return Err(invalid(
            "scripts and event listeners are not allowed in handout XML",
        ));
    }
    Ok(())
}

fn element_name(element: &BytesStart<'_>) -> Result<String> {
    std::str::from_utf8(element.name().as_ref())
        .map(str::to_owned)
        .map_err(|_err| invalid("handout element name is not UTF-8"))
}

fn end_name(element: &quick_xml::events::BytesEnd<'_>) -> Result<String> {
    std::str::from_utf8(element.name().as_ref())
        .map(str::to_owned)
        .map_err(|_err| invalid("handout closing element name is not UTF-8"))
}

fn find_open_end(xml: &str, start: usize, end: usize) -> Result<usize> {
    let fragment = xml
        .get(start..end)
        .ok_or_else(|| invalid("handout root span is not UTF-8 aligned"))?;
    let mut quote = None;
    for (offset, character) in fragment.char_indices() {
        match (quote, character) {
            (None, '\'' | '"') => quote = Some(character),
            (Some(active), value) if active == value => quote = None,
            (None, '>') => return Ok(start + offset + 1),
            _ => {},
        }
    }
    Err(invalid("handout-master open tag is unterminated"))
}

fn split_bom(xml: &str) -> (&str, &str) {
    xml.strip_prefix('\u{feff}')
        .map_or(("", xml), |body| ("\u{feff}", body))
}

fn replace_range(xml: &str, start: usize, end: usize, replacement: &str) -> Result<String> {
    if start > end || end > xml.len() || !xml.is_char_boundary(start) || !xml.is_char_boundary(end)
    {
        return Err(invalid("invalid handout XML replacement span"));
    }
    let mut output = String::with_capacity(xml.len() - (end - start) + replacement.len());
    output.push_str(&xml[..start]);
    output.push_str(replacement);
    output.push_str(&xml[end..]);
    Ok(output)
}

fn insert_into_container(xml: &str, location: &Location, fragment: &str) -> Result<String> {
    if location.empty {
        let raw = &xml[location.start..location.open_end];
        let slash = raw
            .rfind("/>")
            .ok_or_else(|| invalid("empty master-styles marker is missing"))?;
        let replacement = format!(
            "{}>{fragment}</{}>",
            raw[..slash].trim_end(),
            location.qname
        );
        replace_range(xml, location.start, location.end, &replacement)
    } else {
        replace_range(xml, location.content_end, location.content_end, fragment)
    }
}

fn canonical_open(master: &Master) -> String {
    let mut output = format!(
        "<style:handout-master xmlns:style=\"{}\" xmlns:draw=\"{}\" xmlns:presentation=\"{}\" xmlns:office=\"{}\" xmlns:dr3d=\"{}\" xmlns:svg=\"{}\" xmlns:text=\"{}\" xmlns:xlink=\"{}\" style:page-layout-name=\"{}\"",
        String::from_utf8_lossy(STYLE),
        String::from_utf8_lossy(DRAW),
        String::from_utf8_lossy(PRESENTATION),
        String::from_utf8_lossy(OFFICE),
        String::from_utf8_lossy(DR3D),
        String::from_utf8_lossy(SVG),
        String::from_utf8_lossy(TEXT),
        String::from_utf8_lossy(XLINK),
        escape_attr(&master.page_layout_name),
    );
    append_optional(
        &mut output,
        "presentation",
        "presentation-page-layout-name",
        master.presentation_page_layout_name.as_deref(),
    );
    append_optional(
        &mut output,
        "draw",
        "style-name",
        master.drawing_style_name.as_deref(),
    );
    append_optional(
        &mut output,
        "presentation",
        "use-header-name",
        master.header_name.as_deref(),
    );
    append_optional(
        &mut output,
        "presentation",
        "use-footer-name",
        master.footer_name.as_deref(),
    );
    append_optional(
        &mut output,
        "presentation",
        "use-date-time-name",
        master.date_time_name.as_deref(),
    );
    output.push('>');
    output
}

fn append_optional(output: &mut String, prefix: &str, local: &str, value: Option<&str>) {
    if let Some(value) = value {
        output.push(' ');
        output.push_str(prefix);
        output.push(':');
        output.push_str(local);
        output.push_str("=\"");
        output.push_str(&escape_attr(value));
        output.push('"');
    }
}

fn fragment_parts(xml: &str) -> Result<FragmentParts> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut start = None;
    let mut open_end = 0usize;
    let mut qname = String::new();
    loop {
        let event_start = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid(format!("handout fragment parsing error: {error}")))?;
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    start = Some(event_start);
                    open_end = event_end;
                    qname = element_name(&element)?;
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    return Ok(FragmentParts {
                        qname: element_name(&element)?,
                        open: xml.to_string(),
                        inner: String::new(),
                    });
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("handout fragment depth underflow"))?;
                if depth == 0 {
                    let start = start.ok_or_else(|| invalid("handout fragment has no root"))?;
                    return Ok(FragmentParts {
                        qname,
                        open: xml[start..open_end].to_string(),
                        inner: xml[open_end..event_start].to_string(),
                    });
                }
            },
            Event::Text(value)
                if depth == 0 && !value.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid(
                    "non-whitespace text is outside handout fragment root",
                ));
            },
            Event::CData(value)
                if depth == 0 && !value.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid(
                    "non-whitespace CDATA is outside handout fragment root",
                ));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD or PI in handout fragment"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Err(invalid("handout fragment has no complete root"))
}

fn parse_start_attributes(xml: &str) -> Result<(String, Vec<Attribute>)> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let event = reader
        .read_event_into(&mut buffer)
        .map_err(|error| invalid(format!("handout root parsing error: {error}")))?;
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element.into_owned(),
        _ => return Err(invalid("handout source does not start with an element")),
    };
    let qname = element_name(&element)?;
    let mut attributes = Vec::new();
    for raw in element.attributes().with_checks(true) {
        let raw =
            raw.map_err(|error| invalid(format!("invalid handout source attribute: {error}")))?;
        let qname = std::str::from_utf8(raw.key.as_ref())
            .map(str::to_owned)
            .map_err(|_err| invalid("handout source attribute name is not UTF-8"))?;
        let (namespace, local) = reader.resolver().resolve_attribute(raw.key);
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(value)) => Some(value.to_vec()),
            ResolveResult::Unbound => None,
            ResolveResult::Unknown(prefix) => {
                let prefix: &[u8] = prefix.as_ref();
                match prefix {
                    b"office" => Some(OFFICE.to_vec()),
                    b"style" => Some(STYLE.to_vec()),
                    b"draw" => Some(DRAW.to_vec()),
                    b"presentation" => Some(PRESENTATION.to_vec()),
                    b"dr3d" => Some(DR3D.to_vec()),
                    b"svg" => Some(SVG.to_vec()),
                    b"text" => Some(TEXT.to_vec()),
                    b"xlink" => Some(XLINK.to_vec()),
                    _ => None,
                }
            },
        };
        let local = std::str::from_utf8(local.as_ref())
            .map(str::to_owned)
            .map_err(|_err| invalid("handout source attribute local name is not UTF-8"))?;
        let value = raw
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid handout source attribute value: {error}")))?
            .into_owned();
        attributes.push(Attribute {
            qname,
            namespace,
            local,
            value,
        });
    }
    Ok((qname, attributes))
}

fn render_open(source: &str, master: &Master) -> Result<String> {
    let (qname, mut attributes) = parse_start_attributes(source)?;
    let known = [
        (
            STYLE,
            "page-layout-name",
            Some(master.page_layout_name.as_str()),
            "style",
        ),
        (
            PRESENTATION,
            "presentation-page-layout-name",
            master.presentation_page_layout_name.as_deref(),
            "presentation",
        ),
        (
            DRAW,
            "style-name",
            master.drawing_style_name.as_deref(),
            "draw",
        ),
        (
            PRESENTATION,
            "use-header-name",
            master.header_name.as_deref(),
            "presentation",
        ),
        (
            PRESENTATION,
            "use-footer-name",
            master.footer_name.as_deref(),
            "presentation",
        ),
        (
            PRESENTATION,
            "use-date-time-name",
            master.date_time_name.as_deref(),
            "presentation",
        ),
    ];
    let mut required_namespaces = Vec::new();
    for (namespace, local, _value, preferred) in known {
        let prefix = attributes
            .iter()
            .find(|attribute| {
                attribute.namespace.as_deref() == Some(namespace) && attribute.local == local
            })
            .and_then(|attribute| {
                attribute
                    .qname
                    .split_once(':')
                    .map(|(prefix, _)| prefix.to_string())
            })
            .unwrap_or_else(|| preferred.to_string());
        if !attributes.iter().any(|attribute| {
            attribute.qname == format!("xmlns:{prefix}")
                && attribute.value == String::from_utf8_lossy(namespace)
        }) {
            required_namespaces.push((prefix, String::from_utf8_lossy(namespace).into_owned()));
        }
    }
    for (prefix, namespace) in [
        ("office", OFFICE),
        ("style", STYLE),
        ("draw", DRAW),
        ("presentation", PRESENTATION),
        ("dr3d", DR3D),
        ("svg", SVG),
        ("text", TEXT),
        ("xlink", XLINK),
    ] {
        required_namespaces.push((
            prefix.to_string(),
            String::from_utf8_lossy(namespace).into_owned(),
        ));
    }
    for (prefix, namespace) in required_namespaces {
        if !attributes
            .iter()
            .any(|attribute| attribute.qname == format!("xmlns:{prefix}"))
        {
            attributes.push(Attribute {
                qname: format!("xmlns:{prefix}"),
                namespace: None,
                local: format!("xmlns:{prefix}"),
                value: namespace,
            });
        }
    }
    for (namespace, local, value, preferred) in known {
        let position = attributes.iter().position(|attribute| {
            attribute.namespace.as_deref() == Some(namespace) && attribute.local == local
        });
        match (position, value) {
            (Some(position), Some(value)) => attributes[position].value = value.to_string(),
            (Some(position), None) => {
                attributes.remove(position);
            },
            (None, Some(value)) => {
                let prefix = attributes
                    .iter()
                    .find_map(|attribute| {
                        (attribute.qname.starts_with("xmlns:")
                            && attribute.value == String::from_utf8_lossy(namespace))
                        .then(|| attribute.qname[6..].to_string())
                    })
                    .unwrap_or_else(|| preferred.to_string());
                attributes.push(Attribute {
                    qname: format!("{prefix}:{local}"),
                    namespace: Some(namespace.to_vec()),
                    local: local.to_string(),
                    value: value.to_string(),
                });
            },
            (None, None) => {},
        }
    }
    let mut output = String::from("<");
    output.push_str(&qname);
    for attribute in attributes {
        output.push(' ');
        output.push_str(&attribute.qname);
        output.push_str("=\"");
        output.push_str(&escape_attr(&attribute.value));
        output.push('"');
    }
    output.push('>');
    Ok(output)
}

fn escape_attr(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            '\r' => output.push_str("&#13;"),
            '\n' => output.push_str("&#10;"),
            '\t' => output.push_str("&#9;"),
            _ => output.push(character),
        }
    }
    output
}
