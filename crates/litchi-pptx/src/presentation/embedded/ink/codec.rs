use crate::Result;
use crate::presentation::embedded::{
    MAX_XML_DEPTH, increment_nodes, invalid, is_presentationml_name, limit, relationship_value,
    validate_root,
};
use litchi_ooxml_common::process_ooxml;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

pub(crate) const MAX_SLIDE_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_INK_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_TOTAL_INK_BYTES: usize = 256 * 1024 * 1024;
pub(crate) const MAX_ANNOTATIONS: usize = 4_096;
pub(crate) const MAX_CONTENT_PARTS: usize = 4_096;
pub(crate) const MAX_TRACES: usize = 65_536;
pub(crate) const MAX_TRACE_GROUPS: usize = 65_536;
pub(crate) const MAX_RELATIONSHIP_ID_BYTES: usize = 1_024;
const INKML_NAMESPACE: &[u8] = b"http://www.w3.org/2003/InkML";

#[derive(Debug, Default, Clone, Copy)]
pub(crate) struct Summary {
    pub(crate) traces: usize,
    pub(crate) groups: usize,
}

pub(crate) fn scan_slide(xml: &[u8]) -> Result<Vec<String>> {
    if xml.len() > MAX_SLIDE_BYTES {
        return Err(limit("InkML slide XML bytes", MAX_SLIDE_BYTES));
    }
    let xml = process_ooxml(xml)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut ids = Vec::new();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("InkML slide depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("InkML slide depth", MAX_XML_DEPTH));
                }
                if depth == 1 {
                    validate_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                }
                if is_presentationml_name(&namespace, element.name(), b"contentPart") {
                    push_id(&mut ids, &element, decoder, &resolver)?;
                }
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("InkML slide depth", MAX_XML_DEPTH))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("InkML slide depth", MAX_XML_DEPTH));
                }
                if child_depth == 1 {
                    validate_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                    root_closed = true;
                }
                if is_presentationml_name(&namespace, element.name(), b"contentPart") {
                    push_id(&mut ids, &element, decoder, &resolver)?;
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid InkML slide nesting"));
                }
                if depth == 1 {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid("InkML slide must close with p:sld"));
                    }
                    root_closed = true;
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "InkML slide rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => {
                if !root_seen || !root_closed || depth != 0 {
                    return Err(invalid("unterminated InkML slide"));
                }
                return Ok(ids);
            },
            _ => {},
        }
    }
}

fn push_id(
    ids: &mut Vec<String>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<()> {
    if ids.len() >= MAX_CONTENT_PARTS {
        return Err(limit("InkML content-part count", MAX_CONTENT_PARTS));
    }
    let id = relationship_value(element, b"id", decoder, resolver)?
        .ok_or_else(|| invalid("PresentationML contentPart is missing r:id"))?;
    if id.is_empty() || id.len() > MAX_RELATIONSHIP_ID_BYTES {
        return Err(invalid("PresentationML contentPart has an invalid r:id"));
    }
    ids.push(id);
    Ok(())
}

pub(crate) fn inspect(xml: &[u8]) -> Result<Summary> {
    if xml.len() > MAX_INK_BYTES {
        return Err(limit("InkML part bytes", MAX_INK_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut result = Summary::default();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("InkML depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("InkML depth", MAX_XML_DEPTH));
                }
                if depth == 1 {
                    validate_ink_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                }
                observe(&mut result, &namespace, element.name())?;
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("InkML depth", MAX_XML_DEPTH))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("InkML depth", MAX_XML_DEPTH));
                }
                if child_depth == 1 {
                    validate_ink_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                    root_closed = true;
                }
                observe(&mut result, &namespace, element.name())?;
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid InkML nesting"));
                }
                if depth == 1 {
                    if !is_ink_name(&namespace, element.name(), b"ink") {
                        return Err(invalid("InkML must close with ink"));
                    }
                    root_closed = true;
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("InkML rejects DTDs and processing instructions"));
            },
            Event::Eof => {
                if !root_seen || !root_closed || depth != 0 {
                    return Err(invalid("unterminated InkML document"));
                }
                return Ok(result);
            },
            _ => {},
        }
    }
}

fn observe(result: &mut Summary, namespace: &ResolveResult<'_>, name: QName<'_>) -> Result<()> {
    if is_ink_name(namespace, name, b"trace") {
        result.traces = result
            .traces
            .checked_add(1)
            .ok_or_else(|| limit("InkML trace count", MAX_TRACES))?;
        if result.traces > MAX_TRACES {
            return Err(limit("InkML trace count", MAX_TRACES));
        }
    } else if is_ink_name(namespace, name, b"traceGroup") {
        result.groups = result
            .groups
            .checked_add(1)
            .ok_or_else(|| limit("InkML trace-group count", MAX_TRACE_GROUPS))?;
        if result.groups > MAX_TRACE_GROUPS {
            return Err(limit("InkML trace-group count", MAX_TRACE_GROUPS));
        }
    }
    Ok(())
}

fn validate_ink_root(namespace: &ResolveResult<'_>, name: QName<'_>, seen: bool) -> Result<()> {
    if seen || !is_ink_name(namespace, name, b"ink") {
        return Err(invalid("InkML must have one ink root"));
    }
    Ok(())
}

fn is_ink_name(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    name.local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == INKML_NAMESPACE)
}

pub(crate) fn root_namespaces(xml: &[u8]) -> Result<(&'static str, &'static str)> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if !is_presentationml_name(&namespace, element.name(), b"sld") {
                    return Err(invalid("slide root is not PresentationML"));
                }
                let strict = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == crate::presentation::embedded::STRICT_PML);
                return Ok((
                    if strict {
                        "http://purl.oclc.org/ooxml/presentationml/main"
                    } else {
                        "http://schemas.openxmlformats.org/presentationml/2006/main"
                    },
                    if strict {
                        "http://purl.oclc.org/ooxml/officeDocument/relationships"
                    } else {
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    },
                ));
            },
            Event::Eof => return Err(invalid("slide has no root")),
            _ => {},
        }
    }
}
