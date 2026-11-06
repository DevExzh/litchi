use crate::Result;
use crate::presentation::embedded::{
    MAX_XML_DEPTH, increment_nodes, invalid, is_presentationml_name, limit, relationship_value,
    validate_root,
};
use litchi_ooxml_common::mce::{Capabilities, Limits as MceLimits, process_markup_compatibility};
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
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            xml.len(),
            MAX_SLIDE_BYTES,
        ));
    }
    // Native PowerPoint ink is a p:contentPart inside a Choice requiring
    // PowerPoint 2010. This owner understands that inert anchor; using the
    // baseline capabilities instead selects its picture-only fallback.
    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities.understand_namespace("http://schemas.microsoft.com/office/powerpoint/2010/main");
    let limits = MceLimits {
        max_input_bytes: MAX_SLIDE_BYTES,
        max_output_bytes: MAX_SLIDE_BYTES,
        max_depth: MAX_XML_DEPTH,
        ..MceLimits::default()
    };
    let xml = process_markup_compatibility(xml, &capabilities, &limits)?.xml;
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
                depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        MAX_XML_DEPTH,
                    )
                })?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        MAX_XML_DEPTH,
                    ));
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
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        MAX_XML_DEPTH,
                    )
                })?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        child_depth,
                        MAX_XML_DEPTH,
                    ));
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
        return Err(limit(
            crate::error::CodecLimitKind::ContentParts,
            ids.len().checked_add(1).unwrap_or(usize::MAX),
            MAX_CONTENT_PARTS,
        ));
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
        return Err(limit(
            crate::error::CodecLimitKind::PartBytes,
            xml.len(),
            MAX_INK_BYTES,
        ));
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
                depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        MAX_XML_DEPTH,
                    )
                })?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        MAX_XML_DEPTH,
                    ));
                }
                if depth == 1 {
                    validate_ink_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                }
                observe(&mut result, &namespace, element.name())?;
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        MAX_XML_DEPTH,
                    )
                })?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        child_depth,
                        MAX_XML_DEPTH,
                    ));
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
            .ok_or_else(|| limit(crate::error::CodecLimitKind::Traces, usize::MAX, MAX_TRACES))?;
        if result.traces > MAX_TRACES {
            return Err(limit(
                crate::error::CodecLimitKind::Traces,
                result.traces,
                MAX_TRACES,
            ));
        }
    } else if is_ink_name(namespace, name, b"traceGroup") {
        result.groups = result.groups.checked_add(1).ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::TraceGroups,
                usize::MAX,
                MAX_TRACE_GROUPS,
            )
        })?;
        if result.groups > MAX_TRACE_GROUPS {
            return Err(limit(
                crate::error::CodecLimitKind::TraceGroups,
                result.groups,
                MAX_TRACE_GROUPS,
            ));
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

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "codec boundary assertions"
)]
mod tests {
    use super::*;
    fn slide(content: &str) -> String {
        format!(
            r#"<q:sld xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:ext="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:future="urn:future"><q:cSld><q:spTree>{content}</q:spTree></q:cSld></q:sld>"#
        )
    }
    #[test]
    fn ink_choice_uses_namespace_identity_and_skips_inactive_content() {
        let content = r#"<mc:AlternateContent><mc:Choice Requires="future"><q:contentPart rel:id="inactive"/></mc:Choice><mc:Choice Requires="ext"><q:contentPart rel:id="active"/></mc:Choice><mc:Fallback><q:contentPart rel:id="fallback"/></mc:Fallback></mc:AlternateContent>"#;
        assert_eq!(scan_slide(slide(content).as_bytes()).unwrap(), ["active"]);
        let foreign = slide(content).replace(
            "http://schemas.microsoft.com/office/powerpoint/2010/main",
            "urn:other",
        );
        assert_eq!(scan_slide(foreign.as_bytes()).unwrap(), ["fallback"]);
        let no_id = content.replace("rel:id=\"active\"", "");
        assert!(
            matches!(scan_slide(slide(&no_id).as_bytes()),Err(crate::Error::Invalid(message)) if message.contains("missing r:id"))
        );
    }
    #[test]
    fn selected_ink_anchors_keep_the_existing_content_count_limit() {
        let content = "<q:contentPart rel:id=\"same\"/>".repeat(MAX_CONTENT_PARTS + 1);
        assert!(matches!(
            scan_slide(slide(&content).as_bytes()),
            Err(crate::Error::Limit {
                kind: crate::error::LimitKind::Codec(crate::error::CodecLimitKind::ContentParts),
                observed: 4097,
                maximum: 4096
            })
        ));
    }
}
