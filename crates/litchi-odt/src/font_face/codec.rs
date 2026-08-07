//! XML codec for ODF font-face declarations.

use super::model::{
    Declarations, Face, Family, Length, Link, Metric, MetricKind, Pitch, Source, Stretch, Style,
    Variant, Weight, add_text_bytes, validate_text_encoding, validate_value,
};
use super::{
    MAX_DOCUMENT_XML_BYTES, MAX_FONT_FACES, MAX_FORMATS_PER_SOURCE, MAX_SOURCES_PER_FACE,
    MAX_XML_DEPTH, NamespaceKind, OFFICE_NAMESPACE, STYLE_NAMESPACE, SVG_NAMESPACE,
    XLINK_NAMESPACE, invalid, xml_error,
};
use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
    reader::NsReader,
};
use std::collections::HashSet;

pub(super) fn write_declarations(declarations: &Declarations) -> Result<String> {
    declarations.validate()?;
    let mut output = String::with_capacity(256 + declarations.faces.len() * 128);
    output.push_str("<office:font-face-decls xmlns:office=\"");
    output.push_str(std::str::from_utf8(OFFICE_NAMESPACE).expect("namespace is UTF-8"));
    output.push_str("\" xmlns:style=\"");
    output.push_str(std::str::from_utf8(STYLE_NAMESPACE).expect("namespace is UTF-8"));
    output.push_str("\" xmlns:svg=\"");
    output.push_str(std::str::from_utf8(SVG_NAMESPACE).expect("namespace is UTF-8"));
    output.push_str("\" xmlns:xlink=\"");
    output.push_str(std::str::from_utf8(XLINK_NAMESPACE).expect("namespace is UTF-8"));
    output.push_str("\">");
    for face in &declarations.faces {
        output.push_str("<style:font-face");
        write_attr(&mut output, "style:name", Some(&face.name));
        write_attr(
            &mut output,
            "style:font-adornments",
            face.font_adornments.as_deref(),
        );
        write_attr(
            &mut output,
            "style:font-family-generic",
            face.generic_family.map(Family::as_str),
        );
        write_attr(
            &mut output,
            "style:font-pitch",
            face.pitch.map(Pitch::as_str),
        );
        write_attr(&mut output, "style:font-charset", face.charset.as_deref());
        write_attr(&mut output, "svg:font-family", face.family.as_deref());
        write_attr(&mut output, "svg:font-style", face.style.map(Style::as_str));
        write_attr(
            &mut output,
            "svg:font-variant",
            face.variant.map(Variant::as_str),
        );
        write_attr(
            &mut output,
            "svg:font-weight",
            face.weight.map(Weight::as_str),
        );
        write_attr(
            &mut output,
            "svg:font-stretch",
            face.stretch.map(Stretch::as_str),
        );
        write_attr(
            &mut output,
            "svg:font-size",
            face.size.as_ref().map(Length::as_str),
        );
        write_attr(
            &mut output,
            "svg:unicode-range",
            face.unicode_range.as_deref(),
        );
        let mut metrics: Vec<_> = face.metrics.iter().collect();
        metrics.sort_unstable_by_key(|metric| metric.kind.order());
        for metric in metrics {
            output.push_str(" svg:");
            output.push_str(metric.kind.as_str());
            output.push_str("=\"");
            output.push_str(&metric.value.to_string());
            output.push('"');
        }
        write_attr(&mut output, "svg:panose-1", face.panose_1.as_deref());
        write_attr(&mut output, "svg:widths", face.widths.as_deref());
        write_attr(&mut output, "svg:bbox", face.bounding_box.as_deref());

        if face.sources.is_empty() && face.definition_source.is_none() {
            output.push_str("/>");
            continue;
        }
        output.push('>');
        if !face.sources.is_empty() {
            output.push_str("<svg:font-face-src>");
            for source in &face.sources {
                match source {
                    Source::Uri { link, formats } => {
                        output.push_str("<svg:font-face-uri");
                        write_link_attrs(&mut output, link);
                        if formats.is_empty() {
                            output.push_str("/>");
                        } else {
                            output.push('>');
                            for format in formats {
                                output.push_str("<svg:font-face-format");
                                write_attr(&mut output, "svg:string", format.as_deref());
                                output.push_str("/>");
                            }
                            output.push_str("</svg:font-face-uri>");
                        }
                    },
                    Source::LocalName(name) => {
                        output.push_str("<svg:font-face-name");
                        write_attr(&mut output, "svg:name", name.as_deref());
                        output.push_str("/>");
                    },
                }
            }
            output.push_str("</svg:font-face-src>");
        }
        if let Some(link) = &face.definition_source {
            output.push_str("<svg:definition-src");
            write_link_attrs(&mut output, link);
            output.push_str("/>");
        }
        output.push_str("</style:font-face>");
    }
    output.push_str("</office:font-face-decls>");
    Ok(output)
}

/// Parse an optional direct `office:font-face-decls` child from ODF XML.
pub fn parse(xml: &str) -> Result<Option<Declarations>> {
    if xml.len() > MAX_DOCUMENT_XML_BYTES {
        return invalid(format!(
            "ODF XML exceeds the {MAX_DOCUMENT_XML_BYTES} byte font-face limit"
        ));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut result = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"font-face-decls" =>
            {
                if depth != 1 {
                    return invalid("office:font-face-decls must be a direct document child");
                }
                if result.is_some() {
                    return invalid("ODF XML contains duplicate office:font-face-decls");
                }
                result = Some(parse_declarations(&mut reader)?);
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"font-face-decls" =>
            {
                if depth != 1 {
                    return invalid("office:font-face-decls must be a direct document child");
                }
                if result.replace(Declarations::default()).is_some() {
                    return invalid("ODF XML contains duplicate office:font-face-decls");
                }
            },
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF XML depth overflow".to_string()))?;
                if depth > MAX_XML_DEPTH {
                    return invalid(format!("ODF XML exceeds the {MAX_XML_DEPTH} depth limit"));
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid ODF XML depth".to_string()))?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODF font metadata"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(result)
}

fn parse_declarations(reader: &mut NsReader<&[u8]>) -> Result<Declarations> {
    let mut faces = Vec::new();
    let mut names = HashSet::new();
    let mut text_bytes = 0usize;
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Style
                    && element.local_name().as_ref() == b"font-face" =>
            {
                ensure_face_capacity(faces.len())?;
                let mut face = parse_face_attributes(reader, &element, &mut text_bytes)?;
                parse_face_children(reader, &mut face, &mut text_bytes)?;
                if !names.insert(face.name.clone()) {
                    return invalid(format!("duplicate style:font-face name '{}'", face.name));
                }
                faces.push(face);
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Style
                    && element.local_name().as_ref() == b"font-face" =>
            {
                ensure_face_capacity(faces.len())?;
                let face = parse_face_attributes(reader, &element, &mut text_bytes)?;
                if !names.insert(face.name.clone()) {
                    return invalid(format!("duplicate style:font-face name '{}'", face.name));
                }
                faces.push(face);
            },
            Event::End(element)
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"font-face-decls" =>
            {
                break;
            },
            Event::Text(text) => {
                require_whitespace(&text.decode().map_err(xml_error)?, "font-face-decls")?;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) => {
                return invalid("DOCTYPE is not allowed in font-face declarations");
            },
            Event::Eof => return invalid("unterminated office:font-face-decls"),
            _ => return invalid("unsupported child in office:font-face-decls"),
        }
        buffer.clear();
    }
    Ok(Declarations { faces })
}

fn parse_face_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    text_bytes: &mut usize,
) -> Result<Face> {
    let mut face = Face::default();
    let mut name_seen = false;
    let mut metric_kinds = HashSet::new();
    let mut seen = HashSet::new();
    for (namespace, local, value) in attributes(reader, element)? {
        if !seen.insert((namespace, local.clone())) {
            return invalid("duplicate style:font-face attribute");
        }
        *text_bytes = add_text_bytes(*text_bytes, value.len())?;
        match (namespace, local.as_slice()) {
            (NamespaceKind::Style, b"name") => {
                validate_value(&value, "style:name", false)?;
                face.name = value;
                name_seen = true;
            },
            (NamespaceKind::Style, b"font-adornments") => face.font_adornments = Some(value),
            (NamespaceKind::Style, b"font-family-generic") => {
                face.generic_family = Some(Family::parse(&value)?);
            },
            (NamespaceKind::Style, b"font-pitch") => face.pitch = Some(Pitch::parse(&value)?),
            (NamespaceKind::Style, b"font-charset") => {
                validate_text_encoding(&value)?;
                face.charset = Some(value);
            },
            (NamespaceKind::Svg, b"font-family") => face.family = Some(value),
            (NamespaceKind::Svg, b"font-style") => face.style = Some(Style::parse(&value)?),
            (NamespaceKind::Svg, b"font-variant") => face.variant = Some(Variant::parse(&value)?),
            (NamespaceKind::Svg, b"font-weight") => face.weight = Some(Weight::parse(&value)?),
            (NamespaceKind::Svg, b"font-stretch") => face.stretch = Some(Stretch::parse(&value)?),
            (NamespaceKind::Svg, b"font-size") => face.size = Some(Length::new(value)?),
            (NamespaceKind::Svg, b"unicode-range") => face.unicode_range = Some(value),
            (NamespaceKind::Svg, b"panose-1") => face.panose_1 = Some(value),
            (NamespaceKind::Svg, b"widths") => face.widths = Some(value),
            (NamespaceKind::Svg, b"bbox") => face.bounding_box = Some(value),
            (NamespaceKind::Svg, local) => {
                let Some(kind) = MetricKind::from_local(local) else {
                    return invalid("unsupported SVG style:font-face attribute");
                };
                if !metric_kinds.insert(kind) {
                    return invalid(format!("duplicate svg:{} metric", kind.as_str()));
                }
                let value = value.parse::<i64>().map_err(|_| {
                    Error::InvalidFormat(format!("invalid svg:{} integer", kind.as_str()))
                })?;
                face.metrics.push(Metric { kind, value });
            },
            _ => return invalid("style:font-face attribute has an unsupported namespace"),
        }
    }
    if !name_seen {
        return invalid("style:font-face requires style:name");
    }
    Ok(face)
}

fn parse_face_children(
    reader: &mut NsReader<&[u8]>,
    face: &mut Face,
    text_bytes: &mut usize,
) -> Result<()> {
    let mut source_seen = false;
    let mut definition_seen = false;
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-src" =>
            {
                if source_seen || definition_seen {
                    return invalid("svg:font-face-src is duplicate or out of order");
                }
                source_seen = true;
                face.sources = parse_sources(reader, text_bytes)?;
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-src" =>
            {
                return invalid("svg:font-face-src requires at least one source");
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"definition-src" =>
            {
                if definition_seen {
                    return invalid("duplicate svg:definition-src");
                }
                definition_seen = true;
                let link = parse_link(reader, &element)?;
                *text_bytes = add_text_bytes(*text_bytes, link.href.len())?;
                face.definition_source = Some(link);
            },
            Event::Start(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"definition-src" =>
            {
                if definition_seen {
                    return invalid("duplicate svg:definition-src");
                }
                definition_seen = true;
                let link = parse_link(reader, &element)?;
                require_empty(reader, NamespaceKind::Svg, b"definition-src")?;
                *text_bytes = add_text_bytes(*text_bytes, link.href.len())?;
                face.definition_source = Some(link);
            },
            Event::End(element)
                if namespace == NamespaceKind::Style
                    && element.local_name().as_ref() == b"font-face" =>
            {
                break;
            },
            Event::Text(text) => {
                require_whitespace(&text.decode().map_err(xml_error)?, "font-face")?;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated style:font-face"),
            _ => return invalid("unsupported child in style:font-face"),
        }
        buffer.clear();
    }
    Ok(())
}

fn parse_sources(reader: &mut NsReader<&[u8]>, text_bytes: &mut usize) -> Result<Vec<Source>> {
    let mut sources = Vec::new();
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-uri" =>
            {
                ensure_source_capacity(sources.len())?;
                let link = parse_link(reader, &element)?;
                *text_bytes = add_text_bytes(*text_bytes, link.href.len())?;
                let formats = parse_formats(reader, text_bytes)?;
                sources.push(Source::Uri { link, formats });
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-uri" =>
            {
                ensure_source_capacity(sources.len())?;
                let link = parse_link(reader, &element)?;
                *text_bytes = add_text_bytes(*text_bytes, link.href.len())?;
                sources.push(Source::Uri {
                    link,
                    formats: Vec::new(),
                });
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-name" =>
            {
                ensure_source_capacity(sources.len())?;
                let name = optional_single_svg_attribute(reader, &element, b"name")?;
                if let Some(name) = &name {
                    *text_bytes = add_text_bytes(*text_bytes, name.len())?;
                }
                sources.push(Source::LocalName(name));
            },
            Event::Start(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-name" =>
            {
                ensure_source_capacity(sources.len())?;
                let name = optional_single_svg_attribute(reader, &element, b"name")?;
                require_empty(reader, NamespaceKind::Svg, b"font-face-name")?;
                if let Some(name) = &name {
                    *text_bytes = add_text_bytes(*text_bytes, name.len())?;
                }
                sources.push(Source::LocalName(name));
            },
            Event::End(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-src" =>
            {
                break;
            },
            Event::Text(text) => {
                require_whitespace(&text.decode().map_err(xml_error)?, "font-face-src")?;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated svg:font-face-src"),
            _ => return invalid("unsupported child in svg:font-face-src"),
        }
        buffer.clear();
    }
    if sources.is_empty() {
        return invalid("svg:font-face-src requires at least one source");
    }
    Ok(sources)
}

fn parse_formats(
    reader: &mut NsReader<&[u8]>,
    text_bytes: &mut usize,
) -> Result<Vec<Option<String>>> {
    let mut formats = Vec::new();
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Empty(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-format" =>
            {
                if formats.len() >= MAX_FORMATS_PER_SOURCE {
                    return invalid(format!(
                        "font source exceeds the {MAX_FORMATS_PER_SOURCE} format limit"
                    ));
                }
                let value = optional_single_svg_attribute(reader, &element, b"string")?;
                if let Some(value) = &value {
                    *text_bytes = add_text_bytes(*text_bytes, value.len())?;
                }
                formats.push(value);
            },
            Event::Start(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-format" =>
            {
                if formats.len() >= MAX_FORMATS_PER_SOURCE {
                    return invalid(format!(
                        "font source exceeds the {MAX_FORMATS_PER_SOURCE} format limit"
                    ));
                }
                let value = optional_single_svg_attribute(reader, &element, b"string")?;
                require_empty(reader, NamespaceKind::Svg, b"font-face-format")?;
                if let Some(value) = &value {
                    *text_bytes = add_text_bytes(*text_bytes, value.len())?;
                }
                formats.push(value);
            },
            Event::End(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-uri" =>
            {
                break;
            },
            Event::Text(text) => {
                require_whitespace(&text.decode().map_err(xml_error)?, "font-face-uri")?;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated svg:font-face-uri"),
            _ => return invalid("unsupported child in svg:font-face-uri"),
        }
        buffer.clear();
    }
    Ok(formats)
}

fn parse_link(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Link> {
    let mut kind = None;
    let mut href = None;
    let mut actuate = None;
    for (namespace, local, value) in attributes(reader, element)? {
        if namespace != NamespaceKind::Xlink {
            return invalid("font source link attribute has an unsupported namespace");
        }
        let slot = match local.as_slice() {
            b"type" => &mut kind,
            b"href" => &mut href,
            b"actuate" => &mut actuate,
            _ => return invalid("unsupported XLink font source attribute"),
        };
        if slot.replace(value).is_some() {
            return invalid("duplicate XLink font source attribute");
        }
    }
    if kind.as_deref() != Some("simple") {
        return invalid("font source requires xlink:type='simple'");
    }
    if actuate.as_deref().is_some_and(|value| value != "onRequest") {
        return invalid("font source xlink:actuate must be 'onRequest'");
    }
    let href =
        href.ok_or_else(|| Error::InvalidFormat("font source requires xlink:href".to_string()))?;
    validate_value(&href, "xlink:href", false)?;
    Ok(Link {
        href,
        actuate_on_request: actuate.is_some(),
    })
}

fn optional_single_svg_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
) -> Result<Option<String>> {
    let attributes = attributes(reader, element)?;
    if attributes.is_empty() {
        return Ok(None);
    }
    if attributes.len() != 1
        || attributes[0].0 != NamespaceKind::Svg
        || attributes[0].1.as_slice() != expected
    {
        return invalid("font source element contains an unsupported attribute");
    }
    validate_value(&attributes[0].2, "SVG font source attribute", true)?;
    Ok(Some(
        attributes.into_iter().next().expect("one attribute").2,
    ))
}

fn require_empty(
    reader: &mut NsReader<&[u8]>,
    expected_namespace: NamespaceKind,
    expected_local: &[u8],
) -> Result<()> {
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::End(element)
                if namespace == expected_namespace
                    && element.local_name().as_ref() == expected_local =>
            {
                return Ok(());
            },
            Event::Text(text) => {
                require_whitespace(&text.decode().map_err(xml_error)?, "empty font source")?;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated empty font source element"),
            _ => return invalid("font source element must be empty"),
        }
        buffer.clear();
    }
}

fn attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Vec<(NamespaceKind, Vec<u8>, String)>> {
    let mut output = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&namespace);
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        output.push((namespace, local.as_ref().to_vec(), value));
    }
    Ok(output)
}

pub(super) fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE_NAMESPACE => NamespaceKind::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE_NAMESPACE => NamespaceKind::Style,
        ResolveResult::Bound(value) if value.as_ref() == SVG_NAMESPACE => NamespaceKind::Svg,
        ResolveResult::Bound(value) if value.as_ref() == XLINK_NAMESPACE => NamespaceKind::Xlink,
        _ => NamespaceKind::Other,
    }
}

fn ensure_face_capacity(count: usize) -> Result<()> {
    if count >= MAX_FONT_FACES {
        invalid(format!(
            "font-face declarations exceed the {MAX_FONT_FACES} face limit"
        ))
    } else {
        Ok(())
    }
}

fn ensure_source_capacity(count: usize) -> Result<()> {
    if count >= MAX_SOURCES_PER_FACE {
        invalid(format!(
            "font face exceeds the {MAX_SOURCES_PER_FACE} source limit"
        ))
    } else {
        Ok(())
    }
}

fn require_whitespace(value: &str, context: &str) -> Result<()> {
    if value.trim().is_empty() {
        Ok(())
    } else {
        invalid(format!("{context} cannot contain text"))
    }
}

fn write_link_attrs(output: &mut String, link: &Link) {
    output.push_str(" xlink:type=\"simple\"");
    write_attr(output, "xlink:href", Some(&link.href));
    if link.actuate_on_request {
        output.push_str(" xlink:actuate=\"onRequest\"");
    }
}

fn write_attr(output: &mut String, name: &str, value: Option<&str>) {
    let Some(value) = value else { return };
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    escape_attribute(output, value);
    output.push('"');
}

fn escape_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\r' => output.push_str("&#13;"),
            '\n' => output.push_str("&#10;"),
            '\t' => output.push_str("&#9;"),
            _ => output.push(character),
        }
    }
}
