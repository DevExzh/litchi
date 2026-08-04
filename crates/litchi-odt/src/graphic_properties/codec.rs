//! Bounded XML parsing, serialization, and lossless mutation for graphic styles.

use litchi_core::Result;
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{QName, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

use super::model::{Child, ChildKind, Kind, Namespace, Properties, Style, Styles, bad, safe};
use super::{
    DR3D, DR3D_NS, DRAW, DRAW_NS, FO, FO_NS, MAX_ATTRIBUTES, MAX_DEPTH, MAX_EVENTS, MAX_STYLES,
    MAX_TOTAL, MAX_VALUE, MAX_XML, OFFICE, OFFICE_NS, STYLE, STYLE_NS, SVG, SVG_NS, TEXT, TEXT_NS,
    XLINK, XLINK_NS,
};

pub(super) fn validate_child(kind: ChildKind, xml: &str) -> Result<()> {
    if xml.len() > MAX_VALUE {
        return Err(bad("graphic property child is too large"));
    }
    let wrapped = format!(
        r#"<wrapper xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:dr3d="{DR3D_NS}" xmlns:draw="{DRAW_NS}" xmlns:fo="{FO_NS}" xmlns:svg="{SVG_NS}" xmlns:text="{TEXT_NS}" xmlns:xlink="{XLINK_NS}">{xml}</wrapper>"#
    );
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    let mut depth = 0;
    let mut found = false;
    let mut events = 0;
    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(bad("graphic property child has too many events"));
        }
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if depth >= MAX_DEPTH {
                    return Err(bad("graphic property child is too deep"));
                }
                let current = element(&reader, start.name());
                if depth == 1 {
                    if found
                        || current.0 != kind.namespace()
                        || current.1 != kind.local().as_bytes()
                    {
                        return Err(bad("graphic property child has the wrong expanded name"));
                    }
                    found = true
                }
                depth += 1
            },
            Ok(Event::Empty(start)) => {
                let current = element(&reader, start.name());
                if depth == 1 {
                    if found
                        || current.0 != kind.namespace()
                        || current.1 != kind.local().as_bytes()
                    {
                        return Err(bad("graphic property child has the wrong expanded name"));
                    }
                    found = true
                }
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if depth == 1 && !bytes.iter().all(u8::is_ascii_whitespace) {
                    return Err(bad("graphic property child has sibling text"));
                }
            },
            Ok(Event::CData(text)) => {
                let bytes: &[u8] = text.as_ref();
                if depth == 1 && !bytes.iter().all(u8::is_ascii_whitespace) {
                    return Err(bad("graphic property child has sibling text"));
                }
            },
            Ok(Event::End(_)) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| bad("invalid graphic property child"))?
            },
            Ok(Event::Decl(_)) | Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad(
                    "declarations, DTDs, and processing instructions are not allowed in graphic children",
                ));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid graphic property child: {error}"))),
        }
    }
    if !found || depth != 0 {
        return Err(bad("truncated graphic property child"));
    }
    Ok(())
}

fn property_ns(value: ResolveResult<'_>) -> Option<Namespace> {
    match value {
        ResolveResult::Bound(value) if value.as_ref() == DR3D => Some(Namespace::Dr3d),
        ResolveResult::Bound(value) if value.as_ref() == DRAW => Some(Namespace::Draw),
        ResolveResult::Bound(value) if value.as_ref() == FO => Some(Namespace::Fo),
        ResolveResult::Bound(value) if value.as_ref() == STYLE => Some(Namespace::Style),
        ResolveResult::Bound(value) if value.as_ref() == SVG => Some(Namespace::Svg),
        ResolveResult::Bound(value) if value.as_ref() == TEXT => Some(Namespace::Text),
        _ => None,
    }
}
#[derive(Clone, Copy, PartialEq, Eq)]
enum ElementNs {
    Office,
    Style,
    Dr3d,
    Draw,
    Fo,
    Svg,
    Text,
    Xlink,
    Other,
}
impl PartialEq<Namespace> for ElementNs {
    fn eq(&self, other: &Namespace) -> bool {
        matches!(
            (self, other),
            (Self::Dr3d, Namespace::Dr3d)
                | (Self::Draw, Namespace::Draw)
                | (Self::Fo, Namespace::Fo)
                | (Self::Style, Namespace::Style)
                | (Self::Svg, Namespace::Svg)
                | (Self::Text, Namespace::Text)
        )
    }
}
fn ens(value: ResolveResult<'_>) -> ElementNs {
    match value {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE => ElementNs::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE => ElementNs::Style,
        ResolveResult::Bound(value) if value.as_ref() == DR3D => ElementNs::Dr3d,
        ResolveResult::Bound(value) if value.as_ref() == DRAW => ElementNs::Draw,
        ResolveResult::Bound(value) if value.as_ref() == FO => ElementNs::Fo,
        ResolveResult::Bound(value) if value.as_ref() == SVG => ElementNs::Svg,
        ResolveResult::Bound(value) if value.as_ref() == TEXT => ElementNs::Text,
        ResolveResult::Bound(value) if value.as_ref() == XLINK => ElementNs::Xlink,
        _ => ElementNs::Other,
    }
}
fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (ElementNs, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (ens(namespace), local.as_ref().to_vec())
}
#[allow(clippy::type_complexity)]
fn raw_attrs(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Vec<(Option<Namespace>, Vec<u8>, String)>> {
    let mut out = Vec::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| bad(format!("invalid graphic property attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if out.len() >= MAX_ATTRIBUTES {
            return Err(bad("too many graphic property attributes"));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = property_ns(namespace);
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid graphic property value: {error}")))?
            .into_owned();
        safe(&value, "graphic property value", true)?;
        out.push((namespace, local.as_ref().to_vec(), value))
    }
    Ok(out)
}
fn style_header(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    default: bool,
) -> Result<Option<Style>> {
    let mut family = None;
    let mut name = None;
    let mut parent = None;
    for attribute in start.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid graphic style attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if ens(namespace) != ElementNs::Style {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid graphic style value: {error}")))?
            .into_owned();
        match local.as_ref() {
            b"family" => family = Some(value),
            b"name" => name = Some(value),
            b"parent-style-name" => parent = Some(value),
            _ => {},
        }
    }
    if family.as_deref() != Some("graphic") {
        return Ok(None);
    }
    let value = Style {
        name,
        parent_style_name: parent,
        is_default_style: default,
        properties: None,
    };
    value.validate()?;
    Ok(Some(value))
}
fn parse_properties(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Properties> {
    let attrs = raw_attrs(reader, version, start)?;
    let mut properties = Properties::default();
    let mut seen = HashSet::new();
    for (namespace, local, value) in attrs {
        let namespace =
            namespace.ok_or_else(|| bad("unknown style:graphic-properties attribute namespace"))?;
        let kind = Kind::from_expanded(namespace, &local)
            .ok_or_else(|| bad("unknown style:graphic-properties attribute or wrong namespace"))?;
        if !seen.insert(kind) {
            return Err(bad("duplicate style:graphic-properties attribute"));
        }
        properties.set_lexical(kind, &value)?;
    }
    Ok(properties)
}
fn child_kind(current: &(ElementNs, Vec<u8>)) -> Option<ChildKind> {
    match (current.0, current.1.as_slice()) {
        (ElementNs::Text, b"list-style") => Some(ChildKind::ListStyle),
        (ElementNs::Style, b"background-image") => Some(ChildKind::BackgroundImage),
        (ElementNs::Style, b"columns") => Some(ChildKind::Columns),
        _ => None,
    }
}
fn boundary(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid XML event boundary"))
}
struct Active {
    depth: usize,
    style: Style,
    seen: bool,
    property_depth: Option<usize>,
    child: Option<(ChildKind, usize, usize)>,
}
fn push(out: &mut Vec<Style>, style: Style, total: &mut usize) -> Result<()> {
    if out.len() >= MAX_STYLES
        || out.iter().any(|value| {
            value.name == style.name && value.is_default_style == style.is_default_style
        })
    {
        return Err(bad("duplicate or excessive graphic style"));
    }
    *total += style.to_xml_fragment()?.len();
    if *total > MAX_TOTAL {
        return Err(bad("graphic style data is too large"));
    }
    out.push(style);
    Ok(())
}
/// Parse direct graphic-family styles from standard style containers.
pub fn parse_graphic_style_properties(xml: &str) -> Result<Styles> {
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(ElementNs, Vec<u8>)> = Vec::new();
    let mut active: Option<Active> = None;
    let mut out = Vec::new();
    let mut total = 0;
    let mut events = 0;
    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(bad("styles XML has too many events"));
        }
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("styles XML is too deep"));
                }
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|value| {
                    value.0 == ElementNs::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == ElementNs::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        active = Some(Active {
                            depth,
                            style,
                            seen: false,
                            property_depth: None,
                            child: None,
                        })
                    }
                    continue;
                }
                if let Some(value) = active.as_mut() {
                    if value.child.is_some() {
                        continue;
                    }
                    if depth == value.depth + 1
                        && current.0 == ElementNs::Style
                        && current.1 == b"graphic-properties"
                    {
                        if value.seen {
                            return Err(bad("duplicate style:graphic-properties"));
                        }
                        value.seen = true;
                        value.style.properties = Some(parse_properties(&reader, version, &start)?);
                        value.property_depth = Some(depth)
                    } else if current.1 == b"graphic-properties" {
                        return Err(bad(
                            "style:graphic-properties has invalid namespace or parent",
                        ));
                    } else if value.property_depth.is_some_and(|p| depth == p + 1) {
                        let kind = child_kind(&current)
                            .ok_or_else(|| bad("unexpected style:graphic-properties child"))?;
                        if value
                            .style
                            .properties
                            .as_ref()
                            .unwrap()
                            .child(kind)
                            .is_some()
                        {
                            return Err(bad("duplicate graphic property child"));
                        }
                        value.child = Some((kind, depth, begin))
                    } else if value.property_depth.is_some_and(|p| depth > p) {
                        return Err(bad("unexpected style:graphic-properties child"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|value| {
                    value.0 == ElementNs::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == ElementNs::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        push(&mut out, style, &mut total)?
                    }
                    continue;
                }
                if let Some(value) = active.as_mut() {
                    if value.child.is_some() {
                        continue;
                    }
                    if depth == value.depth + 1
                        && current.0 == ElementNs::Style
                        && current.1 == b"graphic-properties"
                    {
                        if value.seen {
                            return Err(bad("duplicate style:graphic-properties"));
                        }
                        value.seen = true;
                        value.style.properties = Some(parse_properties(&reader, version, &start)?)
                    } else if current.1 == b"graphic-properties" {
                        return Err(bad(
                            "style:graphic-properties has invalid namespace or parent",
                        ));
                    } else if value.property_depth.is_some_and(|p| depth == p + 1) {
                        let kind = child_kind(&current)
                            .ok_or_else(|| bad("unexpected style:graphic-properties child"))?;
                        let child = Child::new(kind, &xml[begin..end])?;
                        if value
                            .style
                            .properties
                            .as_mut()
                            .unwrap()
                            .set_child(child)
                            .is_some()
                        {
                            return Err(bad("duplicate graphic property child"));
                        }
                    } else if value.property_depth.is_some_and(|p| depth > p) {
                        return Err(bad("unexpected style:graphic-properties child"));
                    }
                }
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if active
                    .as_ref()
                    .is_some_and(|value| value.property_depth.is_some() && value.child.is_none())
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(bad("unexpected text in style:graphic-properties"));
                }
            },
            Ok(Event::CData(text)) => {
                let bytes: &[u8] = text.as_ref();
                if active
                    .as_ref()
                    .is_some_and(|value| value.property_depth.is_some() && value.child.is_none())
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(bad("unexpected text in style:graphic-properties"));
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let depth = stack.len();
                if let Some(value) = active.as_mut() {
                    if value
                        .child
                        .as_ref()
                        .is_some_and(|(_, child_depth, _)| *child_depth == depth)
                    {
                        let (kind, _, begin) = value.child.take().unwrap();
                        let child = Child::new(kind, &xml[begin..end])?;
                        value.style.properties.as_mut().unwrap().set_child(child);
                    }
                    if value.property_depth == Some(depth) {
                        value.property_depth = None
                    }
                }
                if active.as_ref().is_some_and(|value| value.depth == depth) {
                    push(&mut out, active.take().unwrap().style, &mut total)?
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    if !stack.is_empty() || active.is_some() {
        return Err(bad("truncated styles XML"));
    }
    Ok(Styles { styles: out })
}

#[derive(Default)]
struct Span {
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
}
#[derive(Default)]
struct Target {
    style: Span,
    properties: Option<Span>,
}
fn replace(xml: &str, span: &Span, value: &str) -> String {
    format!("{}{}{}", &xml[..span.start], value, &xml[span.end..])
}
fn expand(xml: &str, span: &Span, value: &str) -> Result<String> {
    let raw = &xml[span.start..span.end];
    let slash = raw.rfind("/>").ok_or_else(|| bad("invalid empty style"))?;
    Ok(replace(
        xml,
        span,
        &format!("{}>{value}</{}>", &raw[..slash], span.qname),
    ))
}
/// Losslessly replace, insert, or remove one existing graphic style property element.
pub fn set_graphic_style_properties_xml(xml: &str, requested: &Style) -> Result<String> {
    requested.validate()?;
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(ElementNs, Vec<u8>)> = Vec::new();
    let mut target_depth = None;
    let mut active: Option<Target> = None;
    let mut found = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|value| {
                    value.0 == ElementNs::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == ElementNs::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target graphic style"));
                        }
                        target_depth = Some(depth);
                        active = Some(Target {
                            style: Span {
                                start: begin,
                                qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                                ..Default::default()
                            },
                            ..Default::default()
                        })
                    }
                } else if target_depth.is_some_and(|d| depth == d + 1)
                    && current.0 == ElementNs::Style
                    && current.1 == b"graphic-properties"
                {
                    let span = Span {
                        start: begin,
                        qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                        ..Default::default()
                    };
                    if active.as_mut().unwrap().properties.replace(span).is_some() {
                        return Err(bad("duplicate style:graphic-properties"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|value| {
                    value.0 == ElementNs::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == ElementNs::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                let span = Span {
                    start: begin,
                    end,
                    end_start: begin,
                    qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                    empty: true,
                };
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target graphic style"));
                        }
                        found = Some(Target {
                            style: span,
                            ..Default::default()
                        })
                    }
                } else if target_depth.is_some_and(|d| depth == d + 1)
                    && current.0 == ElementNs::Style
                    && current.1 == b"graphic-properties"
                    && active.as_mut().unwrap().properties.replace(span).is_some()
                {
                    return Err(bad("duplicate style:graphic-properties"));
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let depth = stack.len();
                if let Some(spans) = active.as_mut() {
                    if spans.properties.as_ref().is_some_and(|span| span.end == 0)
                        && target_depth.is_some_and(|d| depth == d + 1)
                    {
                        let span = spans.properties.as_mut().unwrap();
                        span.end_start = begin;
                        span.end = end
                    }
                    if target_depth == Some(depth) {
                        spans.style.end_start = begin;
                        spans.style.end = end;
                        found = active.take();
                        target_depth = None
                    }
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    let spans = found.ok_or_else(|| bad("target graphic style does not exist"))?;
    let replacement = requested
        .properties
        .as_ref()
        .map(Properties::to_xml_fragment)
        .transpose()?;
    if let Some(properties) = &spans.properties {
        return Ok(replace(
            xml,
            properties,
            replacement.as_deref().unwrap_or(""),
        ));
    }
    let Some(replacement) = replacement else {
        return Ok(xml.to_owned());
    };
    if spans.style.empty {
        return expand(xml, &spans.style, &replacement);
    }
    let mut out = xml.to_owned();
    out.insert_str(spans.style.end_start, &replacement);
    Ok(out)
}
