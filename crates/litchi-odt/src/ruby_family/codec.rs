//! XML parsing and structure-preserving mutation for ODF ruby content.

use litchi_core::Result;
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    reader::NsReader,
};
use std::{collections::HashSet, ops::Range};

use super::{
    MAX_ATTRIBUTES, MAX_BASE, MAX_DEPTH, MAX_EVENTS, MAX_RUBIES, MAX_VALUE, MAX_XML, Ns, bad,
    model::{
        Alignment, Annotation, Annotations, Base, Entry, Position, Properties, Span, Style, Styles,
    },
    ns,
    ruby_inline_specs::{is_hyperlink_child, is_ruby_base_child},
    ruby_range, validate_style_name, validate_text,
};

fn name(value: &[u8]) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_error| bad("invalid UTF-8 XML name"))
}
type Attrs = Vec<(Ns, String, String)>;
fn attributes(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Attrs> {
    let mut out = Vec::new();
    let mut seen = HashSet::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid ruby attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if out.len() >= MAX_ATTRIBUTES {
            return Err(bad("too many ruby attributes"));
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let key = (ns(&resolved), name(local.as_ref())?);
        if !seen.insert(key.clone()) {
            return Err(bad("duplicate expanded ruby attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| bad(format!("invalid ruby attribute value: {error}")))?
            .into_owned();
        validate_text(&value, "ruby attribute value", true)?;
        out.push((key.0, key.1, value));
    }
    Ok(out)
}
fn only_style_ref(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    context: &str,
) -> Result<Option<String>> {
    let mut style = None;
    for (namespace, local, value) in attributes(reader, start)? {
        if namespace == Ns::Text && local == "style-name" && style.is_none() {
            validate_style_name(&value, context)?;
            style = Some(value);
        } else {
            return Err(bad(format!("unsupported {context} attribute")));
        }
    }
    Ok(style)
}
fn require_no_attrs(reader: &NsReader<&[u8]>, start: &BytesStart<'_>, context: &str) -> Result<()> {
    if attributes(reader, start)?.is_empty() {
        Ok(())
    } else {
        Err(bad(format!("{context} has attributes")))
    }
}

fn parse_style_header(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Option<Style>> {
    let attributes = attributes(reader, start)?;
    if !attributes.iter().any(|(namespace, local, value)| {
        namespace == &Ns::Style && local == "family" && value == "ruby"
    }) {
        return Ok(None);
    }
    let mut family = None;
    let mut style_name = None;
    let mut display = None;
    let mut parent = None;
    for (namespace, local, value) in attributes {
        if namespace != Ns::Style {
            return Err(bad("ruby style attribute has wrong namespace"));
        }
        match local.as_str() {
            "family" => family = Some(value),
            "name" => style_name = Some(value),
            "display-name" => display = Some(value),
            "parent-style-name" => parent = Some(value),
            _ => return Err(bad("unsupported ruby style attribute")),
        }
    }
    if family.as_deref() != Some("ruby") {
        return Ok(None);
    }
    let value = Style {
        name: style_name.ok_or_else(|| bad("ruby style requires style:name"))?,
        display_name: display,
        parent_style_name: parent,
        properties: None,
    };
    value.validate()?;
    Ok(Some(value))
}
fn parse_properties(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Properties> {
    let mut value = Properties::default();
    for (namespace, local, lexical) in attributes(reader, start)? {
        if namespace != Ns::Style {
            return Err(bad("ruby property has wrong attribute namespace"));
        }
        match local.as_str() {
            "ruby-position" if value.position.is_none() => {
                value.position = Some(Position::parse(&lexical)?);
            },
            "ruby-align" if value.alignment.is_none() => {
                value.alignment = Some(Alignment::parse(&lexical)?);
            },
            _ => return Err(bad("unknown or duplicate ruby property")),
        }
    }
    Ok(value)
}

struct ActiveStyle {
    depth: usize,
    value: Style,
    property_depth: Option<usize>,
    seen: bool,
}
pub fn parse_ruby_styles(xml: &str) -> Result<Styles> {
    if xml.len() > MAX_XML {
        return Err(bad("ruby styles XML is too large"));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut active: Option<ActiveStyle> = None;
    let mut styles = Vec::new();
    let mut events = 0usize;
    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(bad("too many ruby style events"));
        }
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid ruby styles XML: {error}")))?;
        let namespace = ns(&resolved);
        match event {
            Event::Start(ref start) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("ruby styles XML is too deep"));
                }
                let local = start.local_name().as_ref().to_vec();
                let direct = matches!(stack.last(), Some((Ns::Office, parent)) if parent.as_slice() == b"styles" || parent.as_slice() == b"automatic-styles")
                    && namespace == Ns::Style
                    && local == b"style";
                stack.push((namespace, local.clone()));
                let depth = stack.len();
                if direct {
                    if let Some(value) = parse_style_header(&reader, start)? {
                        active = Some(ActiveStyle {
                            depth,
                            value,
                            property_depth: None,
                            seen: false,
                        });
                    }
                } else if namespace == Ns::Style && local == b"ruby-properties" {
                    let Some(style) = active.as_mut() else {
                        return Err(bad("style:ruby-properties has invalid placement"));
                    };
                    if depth != style.depth + 1 || style.seen {
                        return Err(bad("duplicate or nested style:ruby-properties"));
                    }
                    style.seen = true;
                    style.value.properties = Some(parse_properties(&reader, start)?);
                    style.property_depth = Some(depth);
                } else if active.as_ref().is_some_and(|style| depth > style.depth) {
                    return Err(bad("ruby style has unsupported child"));
                }
            },
            Event::Empty(ref start) => {
                let local = start.local_name().as_ref().to_vec();
                let depth = stack.len() + 1;
                let direct = matches!(stack.last(), Some((Ns::Office, parent)) if parent.as_slice() == b"styles" || parent.as_slice() == b"automatic-styles")
                    && namespace == Ns::Style
                    && local == b"style";
                if direct {
                    if let Some(value) = parse_style_header(&reader, start)? {
                        styles.push(value);
                    }
                } else if namespace == Ns::Style && local == b"ruby-properties" {
                    let Some(style) = active.as_mut() else {
                        return Err(bad("style:ruby-properties has invalid placement"));
                    };
                    if depth != style.depth + 1 || style.seen {
                        return Err(bad("duplicate or nested style:ruby-properties"));
                    }
                    style.seen = true;
                    style.value.properties = Some(parse_properties(&reader, start)?);
                } else if active.as_ref().is_some_and(|style| depth > style.depth) {
                    return Err(bad("ruby style has unsupported child"));
                }
            },
            Event::Text(ref text) if active.is_some() => {
                let bytes: &[u8] = text.as_ref();
                if !bytes.iter().all(u8::is_ascii_whitespace) {
                    return Err(bad("ruby style cannot contain text"));
                }
            },
            Event::CData(_) if active.is_some() => {
                return Err(bad("ruby style cannot contain CDATA"));
            },
            Event::End(_) => {
                let depth = stack.len();
                if active
                    .as_ref()
                    .is_some_and(|style| style.property_depth == Some(depth))
                {
                    active
                        .as_mut()
                        .ok_or_else(|| bad("missing active ruby style"))?
                        .property_depth = None;
                }
                if active.as_ref().is_some_and(|style| style.depth == depth) {
                    styles.push(
                        active
                            .take()
                            .ok_or_else(|| bad("missing completed ruby style"))?
                            .value,
                    );
                }
                stack.pop();
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(bad("DTD and processing instructions are prohibited"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || active.is_some() {
        return Err(bad("truncated ruby styles XML"));
    }
    let value = Styles { styles };
    value.validate()?;
    Ok(value)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ActiveRuby {
    depth: usize,
    start: usize,
    style_name: Option<String>,
    base_depth: Option<usize>,
    base_start: usize,
    base: Option<(usize, usize)>,
    text_depth: Option<usize>,
    text_seen: bool,
    text_style_name: Option<String>,
    text: String,
}
fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid ruby XML event boundary"))
}
pub(crate) fn ruby_parent(parent: Option<&(Ns, Vec<u8>)>) -> bool {
    matches!(parent, Some((Ns::Text, local)) if matches!(local.as_slice(), b"p" | b"h" | b"span" | b"a" | b"meta" | b"meta-field" | b"ruby-base"))
}
pub(super) fn parse_ruby_entries(xml: &str) -> Result<Vec<Entry>> {
    if xml.len() > MAX_XML {
        return Err(bad("ruby XML is too large"));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut active = Vec::<ActiveRuby>::new();
    let mut entries = Vec::new();
    let mut events = 0usize;
    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(bad("too many ruby XML events"));
        }
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid ruby XML: {error}")))?;
        let namespace = ns(&resolved);
        match event {
            Event::Start(ref start) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("ruby XML is too deep"));
                }
                let end = reader.buffer_position() as usize;
                let begin = event_start(xml, end)?;
                let local = start.local_name().as_ref().to_vec();
                let depth = stack.len() + 1;
                let parent = stack.last();
                if matches!(local.as_slice(), b"ruby" | b"ruby-base" | b"ruby-text")
                    && namespace != Ns::Text
                {
                    return Err(bad("ruby element uses wrong namespace"));
                }
                if active.last().is_some_and(|ruby| ruby.text_depth.is_some()) {
                    return Err(bad("text:ruby-text may contain only text"));
                }
                if namespace == Ns::Text && local == b"ruby" {
                    if !ruby_parent(parent) {
                        return Err(bad("text:ruby has invalid placement"));
                    }
                    if entries.len() + active.len() >= MAX_RUBIES {
                        return Err(bad("too many ruby annotations"));
                    }
                    active.push(ActiveRuby {
                        depth,
                        start: begin,
                        style_name: only_style_ref(&reader, start, "ruby style reference")?,
                        base_depth: None,
                        base_start: 0,
                        base: None,
                        text_depth: None,
                        text_seen: false,
                        text_style_name: None,
                        text: String::new(),
                    });
                } else if let Some(ruby) = active.last_mut() {
                    if depth == ruby.depth + 1 {
                        if ruby.base.is_none()
                            && ruby.base_depth.is_none()
                            && namespace == Ns::Text
                            && local == b"ruby-base"
                        {
                            require_no_attrs(&reader, start, "text:ruby-base")?;
                            ruby.base_depth = Some(depth);
                            ruby.base_start = end;
                        } else if ruby.base.is_some()
                            && !ruby.text_seen
                            && namespace == Ns::Text
                            && local == b"ruby-text"
                        {
                            ruby.text_style_name =
                                only_style_ref(&reader, start, "ruby text style reference")?;
                            ruby.text_depth = Some(depth);
                            ruby.text_seen = true;
                        } else {
                            return Err(bad("text:ruby requires ruby-base then ruby-text"));
                        }
                    } else if ruby.base_depth.is_some()
                        && matches!(parent, Some((Ns::Text, p)) if matches!(p.as_slice(), b"ruby-base" | b"span" | b"meta" | b"meta-field"))
                        && !is_ruby_base_child(namespace, &local)
                    {
                        return Err(bad("unsupported text:ruby-base inline child"));
                    } else if ruby.base_depth.is_some()
                        && matches!(parent, Some((Ns::Text, p)) if p.as_slice() == b"a")
                        && !is_hyperlink_child(namespace, &local)
                    {
                        return Err(bad("unsupported hyperlink inline child"));
                    }
                }
                stack.push((namespace, local));
            },
            Event::Empty(ref start) => {
                let end = reader.buffer_position() as usize;
                let local = start.local_name().as_ref().to_vec();
                let depth = stack.len() + 1;
                let parent = stack.last();
                if matches!(local.as_slice(), b"ruby" | b"ruby-base" | b"ruby-text")
                    && namespace != Ns::Text
                {
                    return Err(bad("ruby element uses wrong namespace"));
                }
                if namespace == Ns::Text && local == b"ruby" {
                    return Err(bad("text:ruby requires ruby-base and ruby-text"));
                }
                if let Some(ruby) = active.last_mut() {
                    if ruby.text_depth.is_some() {
                        return Err(bad("text:ruby-text may contain only text"));
                    }
                    if depth == ruby.depth + 1 {
                        if ruby.base.is_none()
                            && ruby.base_depth.is_none()
                            && namespace == Ns::Text
                            && local == b"ruby-base"
                        {
                            require_no_attrs(&reader, start, "text:ruby-base")?;
                            ruby.base = Some((end, end));
                        } else if ruby.base.is_some()
                            && !ruby.text_seen
                            && namespace == Ns::Text
                            && local == b"ruby-text"
                        {
                            ruby.text_style_name =
                                only_style_ref(&reader, start, "ruby text style reference")?;
                            ruby.text_seen = true;
                        } else {
                            return Err(bad("text:ruby requires ruby-base then ruby-text"));
                        }
                    } else if ruby.base_depth.is_some()
                        && matches!(parent, Some((Ns::Text, p)) if matches!(p.as_slice(), b"ruby-base" | b"span" | b"meta" | b"meta-field"))
                        && !is_ruby_base_child(namespace, &local)
                    {
                        return Err(bad("unsupported text:ruby-base inline child"));
                    } else if ruby.base_depth.is_some()
                        && matches!(parent, Some((Ns::Text, p)) if p.as_slice() == b"a")
                        && !is_hyperlink_child(namespace, &local)
                    {
                        return Err(bad("unsupported hyperlink inline child"));
                    }
                }
            },
            Event::Text(ref value)
                if active.last().is_some_and(|ruby| ruby.text_depth.is_some()) =>
            {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| bad(format!("invalid ruby text: {error}")))?;
                let ruby = active
                    .last_mut()
                    .ok_or_else(|| bad("missing active ruby annotation"))?;
                if ruby.text.len() + value.len() > MAX_VALUE {
                    return Err(bad("ruby pronunciation is too large"));
                }
                ruby.text.push_str(&value);
            },
            Event::CData(ref value)
                if active.last().is_some_and(|ruby| ruby.text_depth.is_some()) =>
            {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| bad(format!("invalid ruby CDATA: {error}")))?;
                active
                    .last_mut()
                    .ok_or_else(|| bad("missing active ruby annotation"))?
                    .text
                    .push_str(&value);
            },
            Event::GeneralRef(ref value)
                if active.last().is_some_and(|ruby| ruby.text_depth.is_some()) =>
            {
                active
                    .last_mut()
                    .ok_or_else(|| bad("missing active ruby annotation"))?
                    .text
                    .push_str(&crate::elements::xml::decode_reference(value, "ruby")?);
            },
            Event::End(_) => {
                let end = reader.buffer_position() as usize;
                let begin = event_start(xml, end)?;
                let depth = stack.len();
                let frame = stack
                    .last()
                    .ok_or_else(|| bad("ruby XML depth underflow"))?;
                if let Some(ruby) = active.last_mut() {
                    if ruby.base_depth == Some(depth) {
                        if frame.0 != Ns::Text || frame.1 != b"ruby-base" {
                            return Err(bad("invalid ruby-base end"));
                        }
                        if begin < ruby.base_start || begin - ruby.base_start > MAX_BASE {
                            return Err(bad("ruby base is too large"));
                        }
                        ruby.base = Some((ruby.base_start, begin));
                        ruby.base_depth = None;
                    }
                    if ruby.text_depth == Some(depth) {
                        ruby.text_depth = None;
                    }
                }
                if let Some(ruby) = active.pop_if(|ruby| ruby.depth == depth) {
                    if frame.0 != Ns::Text
                        || frame.1 != b"ruby"
                        || ruby.base.is_none()
                        || !ruby.text_seen
                        || ruby.base_depth.is_some()
                        || ruby.text_depth.is_some()
                    {
                        return Err(bad("text:ruby requires ruby-base then ruby-text"));
                    }
                    let (base_start, base_end) = ruby
                        .base
                        .ok_or_else(|| bad("text:ruby is missing ruby-base"))?;
                    let value = Annotation::new(
                        ruby.style_name,
                        Base {
                            xml: xml[base_start..base_end].to_owned(),
                        },
                        ruby.text,
                        ruby.text_style_name,
                    )?;
                    entries.push(Entry {
                        value,
                        span: Span {
                            start: ruby.start,
                            end,
                        },
                    });
                }
                stack.pop();
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(bad("DTD and processing instructions are prohibited"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || !active.is_empty() {
        return Err(bad("truncated ruby XML"));
    }
    entries.sort_by_key(|entry| entry.span.start);
    Ok(entries)
}

pub fn parse_ruby_annotations(xml: &str) -> Result<Annotations> {
    Ok(Annotations {
        annotations: parse_ruby_entries(xml)?
            .into_iter()
            .map(|entry| entry.value)
            .collect(),
    })
}
pub fn replace_ruby_annotation_xml(xml: &str, index: usize, value: &Annotation) -> Result<String> {
    value.validate()?;
    let entries = parse_ruby_entries(xml)?;
    let span = &entries
        .get(index)
        .ok_or_else(|| bad("ruby annotation index does not exist"))?
        .span;
    Ok(format!(
        "{}{}{}",
        &xml[..span.start],
        value.to_xml_fragment()?,
        &xml[span.end..]
    ))
}
pub fn remove_ruby_annotation_xml(xml: &str, index: usize) -> Result<String> {
    let entries = parse_ruby_entries(xml)?;
    let span = &entries
        .get(index)
        .ok_or_else(|| bad("ruby annotation index does not exist"))?
        .span;
    Ok(format!("{}{}", &xml[..span.start], &xml[span.end..]))
}
pub fn insert_ruby_annotation_xml(
    xml: &str,
    paragraph_index: usize,
    value: &Annotation,
) -> Result<String> {
    value.validate()?;
    parse_ruby_entries(xml)?;
    let fragment = value.to_xml_fragment()?;
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut count = 0usize;
    let mut target = None::<(usize, String)>;
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid ruby insertion XML: {error}")))?;
        let namespace = ns(&resolved);
        match event {
            Event::Start(ref start) => {
                let local = start.local_name().as_ref().to_vec();
                if namespace == Ns::Text && local == b"p" {
                    if count == paragraph_index {
                        target = Some((stack.len() + 1, name(start.name().as_ref())?));
                    }
                    count += 1;
                }
                stack.push((namespace, local));
            },
            Event::Empty(ref start)
                if namespace == Ns::Text && start.local_name().as_ref() == b"p" =>
            {
                let end = reader.buffer_position() as usize;
                let begin = event_start(xml, end)?;
                if count == paragraph_index {
                    let raw = &xml[begin..end];
                    let slash = raw
                        .rfind("/>")
                        .ok_or_else(|| bad("invalid empty paragraph"))?;
                    let qname = name(start.name().as_ref())?;
                    return Ok(format!(
                        "{}{}>{}</{}>{}",
                        &xml[..begin],
                        &raw[..slash],
                        fragment,
                        qname,
                        &xml[end..]
                    ));
                }
                count += 1;
            },
            Event::End(_) => {
                let depth = stack.len();
                if target
                    .as_ref()
                    .is_some_and(|(target_depth, _)| *target_depth == depth)
                {
                    let begin = event_start(xml, reader.buffer_position() as usize)?;
                    return Ok(format!("{}{}{}", &xml[..begin], fragment, &xml[begin..]));
                }
                stack.pop();
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Err(bad("paragraph index does not exist"))
}

/// Replace a UTF-8 range in one paragraph with a ruby annotation.
///
/// `range` is measured over the concatenated values of text, CDATA, and entity
/// reference nodes in the selected `text:p`, in document order. Existing ruby
/// annotations are excluded from that coordinate space. The range must be
/// non-empty and exactly match its base. A `Base::from_text` value may
/// cover adjacent character-data nodes under one eligible parent. A
/// `Base::from_xml_fragment` value may additionally cover a balanced
/// sequence of legal inline elements at one structural depth. Ancestor
/// elements are never split or cloned, and a range cannot cross an existing
/// ruby annotation.
///
/// The mutation is structural only: it does not resolve links, run scripts, or
/// execute macros embedded elsewhere in the document.
pub fn wrap_ruby_annotation_xml(
    xml: &str,
    paragraph_index: usize,
    range: Range<usize>,
    value: &Annotation,
) -> Result<String> {
    if range.start > range.end {
        return Err(bad("ruby text range starts after it ends"));
    }
    if range.is_empty() {
        return Err(bad("ruby text range must be non-empty"));
    }
    value.validate()?;
    parse_ruby_entries(xml)?;
    let fragment = value.to_xml_fragment()?;
    for span in ruby_range::locate_balanced_ruby_ranges(xml, paragraph_index, &range)? {
        if xml[span.start..span.end] == value.base.xml {
            let mut output = String::with_capacity(xml.len() + fragment.len());
            output.push_str(&xml[..span.start]);
            output.push_str(&fragment);
            output.push_str(&xml[span.end..]);
            parse_ruby_entries(&output)?;
            return Ok(output);
        }
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut paragraph_count = 0usize;
    let mut target_depth = None;
    let mut text_offset = 0usize;
    let mut previous_end = 0usize;
    let mut events = 0usize;
    let mut pending = None;

    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(bad("too many ruby range XML events"));
        }
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid ruby range XML: {error}")))?;
        let namespace = ns(&resolved);
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref start) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("ruby range XML is too deep"));
                }
                let local = start.local_name().as_ref().to_vec();
                if namespace == Ns::Text && local == b"p" {
                    if paragraph_count == paragraph_index {
                        target_depth = Some(stack.len() + 1);
                    }
                    paragraph_count = paragraph_count
                        .checked_add(1)
                        .ok_or_else(|| bad("ruby paragraph count overflow"))?;
                }
                stack.push((namespace, local));
            },
            Event::Empty(ref start)
                if namespace == Ns::Text && start.local_name().as_ref() == b"p" =>
            {
                if paragraph_count == paragraph_index {
                    return Err(bad("ruby text range has no text node"));
                }
                paragraph_count = paragraph_count
                    .checked_add(1)
                    .ok_or_else(|| bad("ruby paragraph count overflow"))?;
            },
            Event::Text(ref text) => {
                let content = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| bad(format!("invalid ruby range text: {error}")))?;
                if let Some(output) = ruby_range::collect_ruby_text_node(
                    xml,
                    previous_end..event_end,
                    content.as_ref(),
                    &mut text_offset,
                    &range,
                    &stack,
                    value,
                    &fragment,
                    target_depth,
                    &mut pending,
                )? {
                    parse_ruby_entries(&output)?;
                    return Ok(output);
                }
            },
            Event::CData(ref text) => {
                let content = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| bad(format!("invalid ruby range CDATA: {error}")))?;
                if let Some(output) = ruby_range::collect_ruby_text_node(
                    xml,
                    previous_end..event_end,
                    content.as_ref(),
                    &mut text_offset,
                    &range,
                    &stack,
                    value,
                    &fragment,
                    target_depth,
                    &mut pending,
                )? {
                    parse_ruby_entries(&output)?;
                    return Ok(output);
                }
            },
            Event::GeneralRef(ref reference) => {
                let content = crate::elements::xml::decode_reference(reference, "ruby range")?;
                if let Some(output) = ruby_range::collect_ruby_text_node(
                    xml,
                    previous_end..event_end,
                    &content,
                    &mut text_offset,
                    &range,
                    &stack,
                    value,
                    &fragment,
                    target_depth,
                    &mut pending,
                )? {
                    parse_ruby_entries(&output)?;
                    return Ok(output);
                }
            },
            Event::End(_) => {
                let depth = stack.len();
                if target_depth == Some(depth) {
                    target_depth = None;
                }
                stack
                    .pop()
                    .ok_or_else(|| bad("ruby range XML depth underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(bad("DTD and processing instructions are prohibited"));
            },
            Event::Eof => break,
            _ => {},
        }
        previous_end = event_end;
        buffer.clear();
    }

    if paragraph_index >= paragraph_count {
        return Err(bad("paragraph index does not exist"));
    }
    if range.end > text_offset {
        return Err(bad("ruby text range is out of bounds"));
    }
    Err(bad(
        "ruby text range must fit inside adjacent character-data nodes",
    ))
}

#[derive(Clone)]
enum StyleSite {
    Content(usize),
    Empty(Span, String),
}
fn locate_ruby_style(xml: &str, target_name: &str) -> Result<(Option<Span>, StyleSite)> {
    parse_ruby_styles(xml)?;
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut target = None;
    let mut open = None::<(usize, usize)>;
    let mut site = None;
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid ruby style mutation XML: {error}")))?;
        let namespace = ns(&resolved);
        match event {
            Event::Start(ref start) => {
                let end = reader.buffer_position() as usize;
                let begin = event_start(xml, end)?;
                let local = start.local_name().as_ref().to_vec();
                let depth = stack.len() + 1;
                if namespace == Ns::Style
                    && local == b"style"
                    && matches!(stack.last(), Some((Ns::Office, parent)) if parent == b"styles" || parent == b"automatic-styles")
                    && parse_style_header(&reader, start)?
                        .is_some_and(|style| style.name == target_name)
                {
                    open = Some((depth, begin));
                }
                stack.push((namespace, local));
            },
            Event::Empty(ref start) => {
                let end = reader.buffer_position() as usize;
                let begin = event_start(xml, end)?;
                let local = start.local_name().as_ref().to_vec();
                if namespace == Ns::Style
                    && local == b"style"
                    && matches!(stack.last(), Some((Ns::Office, parent)) if parent == b"styles" || parent == b"automatic-styles")
                    && parse_style_header(&reader, start)?
                        .is_some_and(|style| style.name == target_name)
                {
                    target = Some(Span { start: begin, end });
                }
                if namespace == Ns::Office && local == b"styles" {
                    site = Some(StyleSite::Empty(
                        Span { start: begin, end },
                        name(start.name().as_ref())?,
                    ));
                }
            },
            Event::End(_) => {
                let depth = stack.len();
                let begin = event_start(xml, reader.buffer_position() as usize)?;
                if open.is_some_and(|(d, _)| d == depth) {
                    let (_, start) = open
                        .take()
                        .ok_or_else(|| bad("missing open ruby style span"))?;
                    target = Some(Span {
                        start,
                        end: reader.buffer_position() as usize,
                    });
                }
                if matches!(stack.last(), Some((Ns::Office, local)) if local == b"styles") {
                    site = Some(StyleSite::Content(begin));
                }
                stack.pop();
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok((
        target,
        site.ok_or_else(|| bad("document has no office:styles"))?,
    ))
}
pub fn set_ruby_style_xml(xml: &str, style: &Style) -> Result<String> {
    style.validate()?;
    let (target, site) = locate_ruby_style(xml, &style.name)?;
    let fragment = style.to_xml_fragment()?;
    if let Some(span) = target {
        return Ok(format!(
            "{}{}{}",
            &xml[..span.start],
            fragment,
            &xml[span.end..]
        ));
    }
    match site {
        StyleSite::Content(at) => Ok(format!("{}{}{}", &xml[..at], fragment, &xml[at..])),
        StyleSite::Empty(span, qname) => {
            let raw = &xml[span.start..span.end];
            let slash = raw
                .rfind("/>")
                .ok_or_else(|| bad("invalid empty office:styles"))?;
            Ok(format!(
                "{}{}>{}</{}>{}",
                &xml[..span.start],
                &raw[..slash],
                fragment,
                qname,
                &xml[span.end..]
            ))
        },
    }
}
pub fn remove_ruby_style_xml(xml: &str, name: &str) -> Result<String> {
    validate_style_name(name, "ruby style name")?;
    let (target, _) = locate_ruby_style(xml, name)?;
    let Some(span) = target else {
        return Ok(xml.to_owned());
    };
    Ok(format!("{}{}", &xml[..span.start], &xml[span.end..]))
}
