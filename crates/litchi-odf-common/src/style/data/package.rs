//! Lossless XML mutation and neutral package accessors.

use super::{
    Frame, Kind, MAX_XML_BYTES, NUMBER, OFFICE, Part, Result, STYLE, Section, Style, Styles,
    Version, bad, byte_offset, collect_attributes, decode, direct_style_section, event_start,
    frame_section, invalid, namespace_uri, parse_data_styles_xml, read_document_version,
    validate_name,
};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

#[derive(Clone)]
struct XmlSpan {
    start: usize,
    end: usize,
}

#[derive(Clone)]
struct ContainerSpan {
    section: Section,
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
}

/// Losslessly insert or replace a data style in an existing style container.
///
/// # Errors
///
/// Returns an error when the XML is malformed, lacks the requested container,
/// or the replacement style is not valid for the document version.
pub fn set_data_style_xml(xml: &str, style: &Style) -> Result<String> {
    let version = document_version(xml)?;
    let fragment = style.to_xml_fragment(version)?;
    let (target, container_candidate) = find_style_span(xml, style.section, &style.name)?;
    if let Some(span) = target {
        return Ok(format!(
            "{}{}{}",
            &xml[..span.start],
            fragment,
            &xml[span.end..]
        ));
    }
    let container_span =
        container_candidate.ok_or_else(|| bad("target ODF style container does not exist"))?;
    if container_span.empty {
        let raw = &xml[container_span.start..container_span.end];
        let slash = raw
            .rfind("/>")
            .ok_or_else(|| bad("invalid empty ODF style container"))?;
        let expanded = format!("{}>{}</{}>", &raw[..slash], fragment, container_span.qname);
        return Ok(format!(
            "{}{}{}",
            &xml[..container_span.start],
            expanded,
            &xml[container_span.end..]
        ));
    }
    Ok(format!(
        "{}{}{}",
        &xml[..container_span.end_start],
        fragment,
        &xml[container_span.end_start..]
    ))
}

/// Losslessly remove one named data style from the requested section.
///
/// # Errors
///
/// Returns an error when the XML is malformed or the named style is absent.
pub fn remove_data_style_xml(xml: &str, section: Section, name: &str) -> Result<String> {
    validate_name(name, "style:name")?;
    let (target_candidate, _) = find_style_span(xml, section, name)?;
    let target_span = target_candidate.ok_or_else(|| bad("target data style does not exist"))?;
    Ok(format!(
        "{}{}",
        &xml[..target_span.start],
        &xml[target_span.end..]
    ))
}

/// Parse data styles from the two XML parts of a regular ODF package.
///
/// # Errors
///
/// Returns an error when either supplied XML part is malformed or violates the
/// data-style grammar.
pub fn parse_package(styles_xml: Option<&str>, content_xml: &str) -> Result<Styles> {
    let mut output = Styles::default();
    if let Some(styles) = styles_xml {
        output.append(parse_data_styles_xml(styles, Part::Styles)?)?;
    }
    output.append(parse_data_styles_xml(content_xml, Part::Content)?)?;
    Ok(output)
}

/// Parse data styles from a flat XML document.
///
/// # Errors
///
/// Returns an error when the XML is malformed or violates the data-style
/// grammar.
pub fn parse_flat(xml: &str) -> Result<Styles> {
    parse_data_styles_xml(xml, Part::Flat)
}

fn document_version(xml: &str) -> Result<Version> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("data-style XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut xml_version = XmlVersion::Implicit1_0;
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid data-style XML: {error}")))?
        {
            Event::Decl(ref declaration) => {
                xml_version = declaration
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            },
            Event::Start(ref element) | Event::Empty(ref element) => {
                return read_document_version(&reader, element, xml_version);
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited");
            },
            Event::Eof => return invalid("missing ODF document root"),
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
}

fn find_style_span(
    xml: &str,
    wanted_section: Section,
    wanted_name: &str,
) -> Result<(Option<XmlSpan>, Option<ContainerSpan>)> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("data-style XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut frames = Vec::<Frame>::new();
    let mut active: Option<(usize, usize)> = None;
    let mut target = None;
    let mut container = None;
    let mut container_depth = None;
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid data-style XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                let namespace = namespace_uri(&resolved)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                let end = byte_offset(reader.buffer_position(), "data-style XML")?;
                let start = event_start(xml, end)?;
                let direct = direct_style_section(&frames);
                frames.push(Frame {
                    namespace: namespace.clone(),
                    local: local.clone(),
                });
                let depth = frames.len();
                let current = frames
                    .last()
                    .ok_or_else(|| bad("data-style frame stack underflow"))?;
                if namespace.as_deref() == Some(OFFICE)
                    && frame_section(current) == Some(wanted_section)
                {
                    if container.is_some() || container_depth.is_some() {
                        return invalid("duplicate target ODF style container");
                    }
                    container_depth = Some(depth);
                    container = Some(ContainerSpan {
                        section: wanted_section,
                        start,
                        end: 0,
                        end_start: 0,
                        qname: decode(element.name().as_ref(), "container QName")?,
                        empty: false,
                    });
                }
                if direct == Some(wanted_section)
                    && namespace.as_deref() == Some(NUMBER)
                    && Kind::parse(&local).is_some()
                {
                    let mut aggregate = 0;
                    let attrs = collect_attributes(
                        &reader,
                        element,
                        XmlVersion::Implicit1_0,
                        &mut aggregate,
                    )?;
                    if attrs.iter().any(|attribute| {
                        attribute.namespace.as_deref() == Some(STYLE)
                            && attribute.local == "name"
                            && attribute.value == wanted_name
                    }) {
                        if active.is_some() || target.is_some() {
                            return invalid("duplicate target data style");
                        }
                        active = Some((depth, start));
                    }
                }
            },
            Event::Empty(ref element) => {
                let namespace = namespace_uri(&resolved)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                let end = byte_offset(reader.buffer_position(), "data-style XML")?;
                let start = event_start(xml, end)?;
                let direct = direct_style_section(&frames);
                let current = Frame {
                    namespace: namespace.clone(),
                    local: local.clone(),
                };
                if frame_section(&current) == Some(wanted_section) {
                    if container.is_some() || container_depth.is_some() {
                        return invalid("duplicate target ODF style container");
                    }
                    container = Some(ContainerSpan {
                        section: wanted_section,
                        start,
                        end,
                        end_start: start,
                        qname: decode(element.name().as_ref(), "container QName")?,
                        empty: true,
                    });
                }
                if direct == Some(wanted_section)
                    && namespace.as_deref() == Some(NUMBER)
                    && Kind::parse(&local).is_some()
                {
                    let mut aggregate = 0;
                    let attrs = collect_attributes(
                        &reader,
                        element,
                        XmlVersion::Implicit1_0,
                        &mut aggregate,
                    )?;
                    if attrs.iter().any(|attribute| {
                        attribute.namespace.as_deref() == Some(STYLE)
                            && attribute.local == "name"
                            && attribute.value == wanted_name
                    }) {
                        if target.is_some() {
                            return invalid("duplicate target data style");
                        }
                        target = Some(XmlSpan { start, end });
                    }
                }
            },
            Event::End(_) => {
                let end = byte_offset(reader.buffer_position(), "data-style XML")?;
                let end_start = event_start(xml, end)?;
                let depth = frames.len();
                if active.is_some_and(|(active_depth, _)| active_depth == depth) {
                    let (_, start) = active
                        .take()
                        .ok_or_else(|| bad("data-style target stack underflow"))?;
                    if target.replace(XmlSpan { start, end }).is_some() {
                        return invalid("duplicate target data style");
                    }
                }
                if container_depth == Some(depth) {
                    let value = container
                        .as_mut()
                        .ok_or_else(|| bad("data-style target container is missing"))?;
                    value.end = end;
                    value.end_start = end_start;
                    container_depth = None;
                }
                frames
                    .pop()
                    .ok_or_else(|| bad("data-style element stack underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited");
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if !frames.is_empty() || active.is_some() || container_depth.is_some() {
        return invalid("truncated data-style XML");
    }
    if container
        .as_ref()
        .is_some_and(|value| value.section != wanted_section)
    {
        return invalid("internal data-style section mismatch");
    }
    Ok((target, container))
}
