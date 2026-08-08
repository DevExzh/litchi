//! Bounded ODF data-style XML codec.

use super::tokens::{parse_map, parse_part_node};
use super::*;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

#[derive(Clone)]
pub(crate) struct Attribute {
    pub(crate) namespace: Option<String>,
    pub(crate) local: String,
    pub(crate) value: String,
}

pub(crate) struct NodeBuilder {
    pub(crate) namespace: Option<String>,
    pub(crate) local: String,
    pub(crate) attributes: Vec<Attribute>,
    pub(crate) text: String,
    pub(crate) children: Vec<Node>,
    pub(crate) start: usize,
}

pub(crate) struct Node {
    pub(crate) namespace: Option<String>,
    pub(crate) local: String,
    pub(crate) attributes: Vec<Attribute>,
    pub(crate) text: String,
    pub(crate) children: Vec<Node>,
    pub(crate) raw: String,
}

#[derive(Clone)]
pub(crate) struct Frame {
    pub(crate) namespace: Option<String>,
    pub(crate) local: String,
}

/// Parse direct data styles from both standard style containers in one XML part.
pub fn parse_data_styles_xml(xml: &str, part: Part) -> Result<Styles> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("data-style XML exceeds 64 MiB");
    }
    if !xml.contains("-style") {
        return Ok(Styles::default());
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut frames = Vec::<Frame>::new();
    let mut nodes = Vec::<NodeBuilder>::new();
    let mut output = Styles::default();
    let mut version = Version::V1_2;
    let mut xml_version = XmlVersion::Implicit1_0;
    let mut aggregate = 0usize;
    let mut events = 0usize;

    loop {
        events += 1;
        if events > MAX_EVENTS {
            return invalid("data-style XML has too many events");
        }
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid data-style XML: {error}")))?;
        match event {
            Event::Decl(ref declaration) => {
                xml_version = declaration
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            },
            Event::Start(ref element) => {
                if frames.len() >= MAX_DEPTH {
                    return invalid("data-style XML is too deep");
                }
                let namespace = namespace_uri(&resolved)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                if frames.is_empty() {
                    version = read_document_version(&reader, element, xml_version)?;
                }
                let direct = direct_style_section(&frames);
                reject_spoofed_container(namespace.as_deref(), &local, direct.is_some())?;
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                if !nodes.is_empty()
                    || direct.is_some()
                        && namespace.as_deref() == Some(NUMBER)
                        && Kind::parse(&local).is_some()
                {
                    let attributes =
                        collect_attributes(&reader, element, xml_version, &mut aggregate)?;
                    nodes.push(NodeBuilder {
                        namespace: namespace.clone(),
                        local: local.clone(),
                        attributes,
                        text: String::new(),
                        children: Vec::new(),
                        start,
                    });
                }
                frames.push(Frame { namespace, local });
            },
            Event::Empty(ref element) => {
                let namespace = namespace_uri(&resolved)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                let direct = direct_style_section(&frames);
                reject_spoofed_container(namespace.as_deref(), &local, direct.is_some())?;
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                if !nodes.is_empty() {
                    let node = Node {
                        namespace,
                        local,
                        attributes: collect_attributes(
                            &reader,
                            element,
                            xml_version,
                            &mut aggregate,
                        )?,
                        text: String::new(),
                        children: Vec::new(),
                        raw: xml[start..end].to_string(),
                    };
                    nodes
                        .last_mut()
                        .expect("active data style")
                        .children
                        .push(node);
                } else if let Some(section) = direct
                    && namespace.as_deref() == Some(NUMBER)
                    && Kind::parse(&local).is_some()
                {
                    let node = Node {
                        namespace,
                        local,
                        attributes: collect_attributes(
                            &reader,
                            element,
                            xml_version,
                            &mut aggregate,
                        )?,
                        text: String::new(),
                        children: Vec::new(),
                        raw: xml[start..end].to_string(),
                    };
                    push_style(&mut output, parse_style_node(node, part, section, version)?)?;
                }
            },
            Event::Text(ref text) if !nodes.is_empty() => {
                let decoded = text
                    .decode()
                    .map_err(|error| bad(format!("invalid data-style text: {error}")))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| bad(format!("invalid data-style entity: {error}")))?;
                add_text(
                    &mut nodes.last_mut().expect("active node").text,
                    &unescaped,
                    &mut aggregate,
                )?;
            },
            Event::CData(ref text) if !nodes.is_empty() => {
                let decoded = reader
                    .decoder()
                    .decode(text.as_ref())
                    .map_err(|error| bad(format!("invalid data-style CDATA: {error}")))?;
                add_text(
                    &mut nodes.last_mut().expect("active node").text,
                    &decoded,
                    &mut aggregate,
                )?;
            },
            Event::GeneralRef(_) if !nodes.is_empty() => {
                return invalid("entity references are prohibited in data styles");
            },
            Event::End(_) => {
                let end = reader.buffer_position() as usize;
                if !nodes.is_empty()
                    && nodes.len() == frames.len() - active_base_depth(&frames, &nodes)
                {
                    let builder = nodes.pop().expect("active data-style node");
                    let node = Node {
                        namespace: builder.namespace,
                        local: builder.local,
                        attributes: builder.attributes,
                        text: builder.text,
                        children: builder.children,
                        raw: xml[builder.start..end].to_string(),
                    };
                    if let Some(parent) = nodes.last_mut() {
                        parent.children.push(node);
                    } else {
                        let section = direct_parent_section(&frames)
                            .ok_or_else(|| bad("misplaced data style"))?;
                        push_style(&mut output, parse_style_node(node, part, section, version)?)?;
                    }
                }
                frames
                    .pop()
                    .ok_or_else(|| bad("data-style element stack underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !frames.is_empty() || !nodes.is_empty() {
        return invalid("truncated data-style XML");
    }
    Ok(output)
}

// The active node stack is always a suffix of the document frame stack. This
// helper avoids storing a second absolute depth in every node.
pub(crate) fn active_base_depth(frames: &[Frame], nodes: &[NodeBuilder]) -> usize {
    frames.len().saturating_sub(nodes.len())
}

pub(crate) fn direct_parent_section(frames: &[Frame]) -> Option<Section> {
    frames
        .get(frames.len().checked_sub(2)?)
        .and_then(frame_section)
}

pub(crate) fn direct_style_section(frames: &[Frame]) -> Option<Section> {
    frames.last().and_then(frame_section)
}

pub(crate) fn frame_section(frame: &Frame) -> Option<Section> {
    if frame.namespace.as_deref() != Some(OFFICE) {
        return None;
    }
    match frame.local.as_str() {
        "styles" => Some(Section::Styles),
        "automatic-styles" => Some(Section::AutomaticStyles),
        _ => None,
    }
}

pub(crate) fn parse_style_node(
    mut node: Node,
    part: Part,
    section: Section,
    version: Version,
) -> Result<Style> {
    if node.namespace.as_deref() != Some(NUMBER) {
        return invalid("data style uses the wrong namespace");
    }
    let kind = Kind::parse(&node.local).ok_or_else(|| bad("unknown data-style container"))?;
    ensure_whitespace(&node.text, "data-style container")?;
    let name = required(&mut node.attributes, STYLE, "name")?;
    let display_name = take(&mut node.attributes, STYLE, "display-name");
    let locale = parse_locale(&mut node.attributes)?;
    let title = take(&mut node.attributes, NUMBER, "title");
    let volatile = take_bool(&mut node.attributes, STYLE, "volatile")?;
    let transliteration_format = take(&mut node.attributes, NUMBER, "transliteration-format");
    let transliteration_language = take(&mut node.attributes, NUMBER, "transliteration-language");
    let transliteration_country = take(&mut node.attributes, NUMBER, "transliteration-country");
    let transliteration_style = take(&mut node.attributes, NUMBER, "transliteration-style")
        .map(|value| TransliterationStyle::parse(&value))
        .transpose()?;
    let automatic_order = take_bool(&mut node.attributes, NUMBER, "automatic-order")?;
    let format_source = take(&mut node.attributes, NUMBER, "format-source")
        .map(|value| FormatSource::parse(&value))
        .transpose()?;
    let truncate_on_overflow = take_bool(&mut node.attributes, NUMBER, "truncate-on-overflow")?;
    reject_remaining(&node.attributes, "data-style container")?;

    let mut text_properties = None;
    let mut parts = Vec::new();
    let mut maps = Vec::new();
    let mut compatibility_alias = false;
    let mut phase = 0u8;
    for child in node.children {
        if child.namespace.as_deref() == Some(STYLE) && child.local == "text-properties" {
            if phase != 0 || text_properties.is_some() {
                return invalid("style:text-properties must be the first and only such child");
            }
            ensure_whitespace(&child.text, "style:text-properties")?;
            if !child.children.is_empty() {
                return invalid("style:text-properties cannot contain elements");
            }
            text_properties = Some(TextProperties { xml: child.raw });
            continue;
        }
        phase = phase.max(1);
        if child.namespace.as_deref() == Some(STYLE) && child.local == "map" {
            phase = 2;
            if maps.len() >= MAX_MAPS {
                return invalid("too many style:map elements");
            }
            maps.push(parse_map(child)?);
        } else {
            if phase == 2 {
                return invalid("formatting tokens cannot follow style:map");
            }
            if parts.len() >= MAX_PARTS {
                return invalid("too many data-style tokens");
            }
            let (token, used_alias) = parse_part_node(child, version)?;
            compatibility_alias |= used_alias;
            parts.push(token);
        }
    }
    let style = Style {
        source_part: part,
        section,
        source_version: version,
        kind,
        name,
        display_name,
        locale,
        title,
        volatile,
        transliteration_format,
        transliteration_language,
        transliteration_country,
        transliteration_style,
        automatic_order,
        format_source,
        truncate_on_overflow,
        text_properties,
        parts,
        maps,
    };
    style.validate_inner(version, compatibility_alias)?;
    Ok(style)
}

pub(crate) fn push_style(output: &mut Styles, style: Style) -> Result<()> {
    if output.styles.len() >= MAX_STYLES {
        return invalid("too many data styles");
    }
    if output
        .get(style.source_part, style.section, &style.name)
        .is_some()
    {
        return invalid("duplicate data style identity");
    }
    output.styles.push(style);
    Ok(())
}

pub(crate) fn collect_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    version: XmlVersion,
    aggregate: &mut usize,
) -> Result<Vec<Attribute>> {
    let mut output = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid data-style attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if output.len() >= MAX_ATTRIBUTES {
            return invalid("data-style element has too many attributes");
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_uri(&resolved)?;
        let local = decode(local.as_ref(), "attribute name")?;
        if !seen.insert((namespace.clone(), local.clone())) {
            return invalid("duplicate expanded data-style attribute");
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid data-style attribute value: {error}")))?
            .into_owned();
        add_size(value.len(), aggregate)?;
        output.push(Attribute {
            namespace,
            local,
            value,
        });
    }
    Ok(output)
}

pub(crate) fn read_document_version(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    version: XmlVersion,
) -> Result<Version> {
    let mut aggregate = 0usize;
    let attributes = collect_attributes(reader, element, version, &mut aggregate)?;
    match attributes
        .iter()
        .find(|attribute| {
            attribute.namespace.as_deref() == Some(OFFICE) && attribute.local == "version"
        })
        .map(|attribute| attribute.value.as_str())
    {
        None | Some("1.0" | "1.1" | "1.2") => Ok(Version::V1_2),
        Some("1.3") => Ok(Version::V1_3),
        Some(value) => invalid(format!("unsupported ODF data-style version '{value}'")),
    }
}

pub(crate) fn parse_locale(attributes: &mut Vec<Attribute>) -> Result<Locale> {
    Ok(Locale {
        language: take(attributes, NUMBER, "language"),
        country: take(attributes, NUMBER, "country"),
        script: take(attributes, NUMBER, "script"),
        rfc_language_tag: take(attributes, NUMBER, "rfc-language-tag"),
    })
}

pub(crate) fn take(
    attributes: &mut Vec<Attribute>,
    namespace: &str,
    local: &str,
) -> Option<String> {
    attributes
        .iter()
        .position(|attribute| {
            attribute.namespace.as_deref() == Some(namespace) && attribute.local == local
        })
        .map(|index| attributes.remove(index).value)
}

pub(crate) fn required(
    attributes: &mut Vec<Attribute>,
    namespace: &str,
    local: &str,
) -> Result<String> {
    take(attributes, namespace, local)
        .ok_or_else(|| bad(format!("missing required {namespace}:{local} attribute")))
}

pub(crate) fn take_bool(
    attributes: &mut Vec<Attribute>,
    namespace: &str,
    local: &str,
) -> Result<Option<bool>> {
    take(attributes, namespace, local)
        .map(|value| parse_bool(&value))
        .transpose()
}

pub(crate) fn take_i64(
    attributes: &mut Vec<Attribute>,
    namespace: &str,
    local: &str,
) -> Result<Option<i64>> {
    take(attributes, namespace, local)
        .map(|value| parse_i64(&value, local))
        .transpose()
}

pub(crate) fn required_i64(
    attributes: &mut Vec<Attribute>,
    namespace: &str,
    local: &str,
) -> Result<i64> {
    let value = required(attributes, namespace, local)?;
    parse_i64(&value, local)
}

pub(crate) fn take_f64(
    attributes: &mut Vec<Attribute>,
    namespace: &str,
    local: &str,
) -> Result<Option<f64>> {
    take(attributes, namespace, local)
        .map(|value| match value.as_str() {
            "INF" => Ok(f64::INFINITY),
            "-INF" => Ok(f64::NEG_INFINITY),
            "NaN" => Ok(f64::NAN),
            _ => {
                let parsed: f64 = value
                    .parse()
                    .map_err(|_| bad(format!("invalid {local} double '{value}'")))?;
                if !parsed.is_finite() {
                    return invalid(format!("invalid {local} double '{value}'"));
                }
                Ok(parsed)
            },
        })
        .transpose()
}

pub(crate) fn take_versioned_i64(
    attributes: &mut Vec<Attribute>,
    local: &str,
    version: Version,
    alias: &mut bool,
) -> Result<Option<i64>> {
    let standard = take(attributes, NUMBER, local);
    let extension = take(attributes, LOEXT, local);
    if standard.is_some() && extension.is_some() {
        return invalid(format!("duplicate standard/LO alias for {local}"));
    }
    if standard.is_some() && version == Version::V1_2 {
        return invalid(format!("number:{local} requires ODF 1.3"));
    }
    if extension.is_some() {
        *alias = true;
    }
    standard
        .or(extension)
        .map(|value| parse_i64(&value, local))
        .transpose()
}

pub(crate) fn take_versioned_u64(
    attributes: &mut Vec<Attribute>,
    local: &str,
    version: Version,
    alias: &mut bool,
) -> Result<Option<u64>> {
    let standard = take(attributes, NUMBER, local);
    let extension = take(attributes, LOEXT, local);
    if standard.is_some() && extension.is_some() {
        return invalid(format!("duplicate standard/LO alias for {local}"));
    }
    if standard.is_some() && version == Version::V1_2 {
        return invalid(format!("number:{local} requires ODF 1.3"));
    }
    if extension.is_some() {
        *alias = true;
    }
    standard
        .or(extension)
        .map(|value| parse_u64(&value, local))
        .transpose()
}

pub(crate) fn take_versioned_bool(
    attributes: &mut Vec<Attribute>,
    local: &str,
    version: Version,
    alias: &mut bool,
) -> Result<Option<bool>> {
    let standard = take(attributes, NUMBER, local);
    let extension = take(attributes, LOEXT, local);
    if standard.is_some() && extension.is_some() {
        return invalid(format!("duplicate standard/LO alias for {local}"));
    }
    if standard.is_some() && version == Version::V1_2 {
        return invalid(format!("number:{local} requires ODF 1.3"));
    }
    if extension.is_some() {
        *alias = true;
    }
    standard
        .or(extension)
        .map(|value| parse_bool(&value))
        .transpose()
}

pub(crate) fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid ODF boolean '{value}'")),
    }
}

pub(crate) fn parse_i64(value: &str, name: &str) -> Result<i64> {
    value
        .parse()
        .map_err(|_| bad(format!("invalid {name} integer '{value}'")))
}

pub(crate) fn parse_u64(value: &str, name: &str) -> Result<u64> {
    let parsed: u64 = value
        .parse()
        .map_err(|_| bad(format!("invalid {name} positive integer '{value}'")))?;
    if parsed == 0 {
        return invalid(format!("invalid {name} positive integer '{value}'"));
    }
    Ok(parsed)
}

pub(crate) fn reject_remaining(attributes: &[Attribute], element: &str) -> Result<()> {
    if let Some(attribute) = attributes.first() {
        return invalid(format!(
            "unexpected {element} attribute {}:{}",
            attribute.namespace.as_deref().unwrap_or(""),
            attribute.local
        ));
    }
    Ok(())
}

pub(crate) fn ensure_no_children(node: &Node, element: &str) -> Result<()> {
    if !node.children.is_empty() {
        return invalid(format!("{element} cannot contain elements"));
    }
    Ok(())
}

pub(crate) fn ensure_empty_node(node: &Node, element: &str) -> Result<()> {
    ensure_no_children(node, element)?;
    ensure_whitespace(&node.text, element)
}

pub(crate) fn ensure_whitespace(value: &str, element: &str) -> Result<()> {
    if value.chars().any(|character| !character.is_whitespace()) {
        return invalid(format!("{element} cannot contain text"));
    }
    Ok(())
}

pub(crate) fn validate_name(value: &str, name: &str) -> Result<()> {
    validate_text(value, name)?;
    if value.is_empty() || value.chars().any(char::is_control) {
        return invalid(format!("invalid {name}"));
    }
    Ok(())
}

pub(crate) fn validate_cell_address(value: &str) -> Result<()> {
    validate_text(value, "style:base-cell-address")?;
    if value.is_empty() || value.chars().any(char::is_whitespace) || !value.contains('.') {
        return invalid("invalid style:base-cell-address");
    }
    Ok(())
}

pub(crate) fn validate_locale(locale: &Locale) -> Result<()> {
    for (value, name) in [
        (locale.language.as_deref(), "number:language"),
        (locale.country.as_deref(), "number:country"),
        (locale.script.as_deref(), "number:script"),
        (
            locale.rfc_language_tag.as_deref(),
            "number:rfc-language-tag",
        ),
    ] {
        if let Some(value) = value {
            validate_text(value, name)?;
            if value.is_empty()
                || value
                    .bytes()
                    .any(|byte| !(byte.is_ascii_alphanumeric() || byte == b'-'))
            {
                return invalid(format!("invalid {name} '{value}'"));
            }
        }
    }
    Ok(())
}

pub(crate) fn validate_optional_string(value: Option<&str>, name: &str) -> Result<()> {
    if let Some(value) = value {
        validate_text(value, name)?;
    }
    Ok(())
}

pub(crate) fn validate_text(value: &str, name: &str) -> Result<()> {
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{name} exceeds 64 KiB"));
    }
    if value.chars().any(
        |character| matches!(character, '\u{0}'..='\u{8}' | '\u{B}' | '\u{C}' | '\u{E}'..='\u{1F}'),
    ) {
        return invalid(format!("{name} contains an invalid XML character"));
    }
    Ok(())
}

pub(crate) fn reject_spoofed_container(
    namespace: Option<&str>,
    local: &str,
    direct: bool,
) -> Result<()> {
    if direct && Kind::parse(local).is_some() && namespace != Some(NUMBER) {
        return invalid("data-style container uses the wrong namespace");
    }
    Ok(())
}

pub(crate) fn add_text(target: &mut String, value: &str, aggregate: &mut usize) -> Result<()> {
    add_size(value.len(), aggregate)?;
    if target.len() + value.len() > MAX_VALUE_BYTES {
        return invalid("data-style text exceeds 64 KiB");
    }
    target.push_str(value);
    Ok(())
}

pub(crate) fn add_size(size: usize, aggregate: &mut usize) -> Result<()> {
    if size > MAX_VALUE_BYTES {
        return invalid("data-style value exceeds 64 KiB");
    }
    *aggregate = aggregate
        .checked_add(size)
        .ok_or_else(|| bad("data-style aggregate size overflow"))?;
    if *aggregate > MAX_AGGREGATE_BYTES {
        return invalid("data-style metadata exceeds 32 MiB");
    }
    Ok(())
}

pub(crate) fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid data-style XML event boundary"))
}

pub(crate) fn namespace_uri(result: &ResolveResult<'_>) -> Result<Option<String>> {
    match result {
        ResolveResult::Bound(Namespace(value)) => Ok(Some(decode(value, "namespace URI")?)),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        )),
    }
}

pub(crate) fn decode(value: &[u8], description: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| bad(format!("invalid UTF-8 {description}")))
}
