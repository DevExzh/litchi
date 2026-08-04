//! Lossless XML mutation and neutral package accessors.

use super::*;
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
pub fn set_data_style_xml(xml: &str, style: &Style) -> Result<String> {
    let version = document_version(xml)?;
    let fragment = style.to_xml_fragment(version)?;
    let (target, container) = find_style_span(xml, style.section, &style.name)?;
    if let Some(span) = target {
        return Ok(format!(
            "{}{}{}",
            &xml[..span.start],
            fragment,
            &xml[span.end..]
        ));
    }
    let container = container.ok_or_else(|| bad("target ODF style container does not exist"))?;
    if container.empty {
        let raw = &xml[container.start..container.end];
        let slash = raw
            .rfind("/>")
            .ok_or_else(|| bad("invalid empty ODF style container"))?;
        let expanded = format!("{}>{}</{}>", &raw[..slash], fragment, container.qname);
        return Ok(format!(
            "{}{}{}",
            &xml[..container.start],
            expanded,
            &xml[container.end..]
        ));
    }
    Ok(format!(
        "{}{}{}",
        &xml[..container.end_start],
        fragment,
        &xml[container.end_start..]
    ))
}

/// Losslessly remove one named data style from the requested section.
pub fn remove_data_style_xml(xml: &str, section: Section, name: &str) -> Result<String> {
    validate_name(name, "style:name")?;
    let (target, _) = find_style_span(xml, section, name)?;
    let target = target.ok_or_else(|| bad("target data style does not exist"))?;
    Ok(format!("{}{}", &xml[..target.start], &xml[target.end..]))
}

/// Parse data styles from the two XML parts of a regular ODF package.
pub fn parse_package(styles_xml: Option<&str>, content_xml: &str) -> Result<Styles> {
    let mut output = Styles::default();
    if let Some(styles) = styles_xml {
        output.append(parse_data_styles_xml(styles, Part::Styles)?)?;
    }
    output.append(parse_data_styles_xml(content_xml, Part::Content)?)?;
    Ok(output)
}

/// Parse data styles from a flat XML document.
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
            _ => {},
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
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let direct = direct_style_section(&frames);
                frames.push(Frame {
                    namespace: namespace.clone(),
                    local: local.clone(),
                });
                let depth = frames.len();
                if namespace.as_deref() == Some(OFFICE)
                    && frame_section(frames.last().expect("pushed frame")) == Some(wanted_section)
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
                let end = reader.buffer_position() as usize;
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
                let end = reader.buffer_position() as usize;
                let end_start = event_start(xml, end)?;
                let depth = frames.len();
                if active.is_some_and(|(active_depth, _)| active_depth == depth) {
                    let (_, start) = active.take().expect("active target");
                    if target.replace(XmlSpan { start, end }).is_some() {
                        return invalid("duplicate target data style");
                    }
                }
                if container_depth == Some(depth) {
                    let value = container.as_mut().expect("active target container");
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
            _ => {},
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
