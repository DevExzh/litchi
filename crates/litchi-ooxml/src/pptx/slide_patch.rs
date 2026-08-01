//! Shared slide-XML patching helpers for PowerPoint extension storage
//! (laser traces, slide-show events, and similar `p:ext` payloads).
//!
//! All mutation is byte-level and bounded: fragments are inserted into the
//! slide's extension list (`p:extLst`) — patched into an existing list,
//! expanding an empty one, or creating one before the slide end tag — while
//! the slide's namespace dialect (transitional or Strict) is preserved.

use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::is_presentationml_name;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 128;

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(what: &str) -> OoxmlError {
    invalid(format!("{what} exceeds the supported safety limit"))
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}

/// The slide's PresentationML namespace URI (transitional or Strict).
pub(crate) fn slide_dialect(xml: &[u8]) -> Result<&'static str> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        match reader.read_event().map_err(xml_error)? {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                return Ok(match namespace {
                    ResolveResult::Bound(Namespace(value))
                        if value == crate::pptx::namespace::STRICT_PRESENTATIONML_NAMESPACE =>
                    {
                        "http://purl.oclc.org/ooxml/presentationml/main"
                    },
                    _ => "http://schemas.openxmlformats.org/presentationml/2006/main",
                });
            },
            Event::Eof => return Err(invalid("slide XML has no root element")),
            _ => {},
        }
    }
}

/// Insert a fragment into the slide's extension list (`p:extLst`).
///
/// The fragment is appended to an existing list, an empty `<p:extLst/>`
/// element is expanded around it, or a new list is created directly before
/// the slide end tag. The slide root must be a single PresentationML `sld`
/// element; DTDs and processing instructions are rejected.
pub(crate) fn insert_extension_fragment(xml: &[u8], fragment: &str) -> Result<Vec<u8>> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut ext_lst_depth = None;
    let mut ext_lst_end = None;
    let mut empty_ext_lst: Option<(usize, usize)> = None;
    let mut root_end = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("slide XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("XML node count"))?;
                if nodes > MAX_XML_NODES {
                    return Err(limit("XML node count"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML depth"))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth"));
                }
                if depth == 1 {
                    if root_seen || !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid(
                            "slide XML must have one PresentationML sld root element",
                        ));
                    }
                    root_seen = true;
                }
                if depth == 2
                    && is_presentationml_name(&namespace, element.name(), b"extLst")
                    && ext_lst_depth.replace(depth).is_some()
                {
                    return Err(invalid("slide has multiple extension lists"));
                }
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("XML node count"))?;
                if nodes > MAX_XML_NODES {
                    return Err(limit("XML node count"));
                }
                if !root_seen {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid(
                            "slide XML must have one PresentationML sld root element",
                        ));
                    }
                    root_seen = true;
                }
                if depth + 1 == 2 && is_presentationml_name(&namespace, element.name(), b"extLst") {
                    // Empty `<p:extLst/>`: remember its range for expansion.
                    if ext_lst_depth.replace(usize::MAX).is_some() {
                        return Err(invalid("slide has multiple extension lists"));
                    }
                    empty_ext_lst = Some((
                        start,
                        usize::try_from(reader.buffer_position())
                            .map_err(|_| invalid("slide XML offset overflow"))?,
                    ));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid slide XML nesting"));
                }
                if ext_lst_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"extLst")
                {
                    ext_lst_end = Some(start);
                }
                if depth == 1 {
                    root_end = Some(start);
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || !root_seen {
        return Err(invalid("unterminated or missing PresentationML slide root"));
    }
    let root_end = root_end.ok_or_else(|| invalid("slide is missing its end tag"))?;

    if let Some((start, end)) = empty_ext_lst {
        // Empty `<p:extLst/>`: replace it with an expanded element wrapping
        // the fragment.
        let element = xml[start..end].trim_ascii_end();
        let open = element
            .strip_suffix(b"/>")
            .ok_or_else(|| invalid("slide extension list is not an empty element"))?;
        let mut output = Vec::with_capacity(xml.len() + fragment.len() + 16);
        output.extend_from_slice(&xml[..start]);
        output.extend_from_slice(open);
        output.extend_from_slice(b">");
        output.extend_from_slice(fragment.as_bytes());
        output.extend_from_slice(b"</p:extLst>");
        output.extend_from_slice(&xml[end..]);
        return Ok(output);
    }
    if let Some(position) = ext_lst_end {
        let mut output = Vec::with_capacity(xml.len() + fragment.len());
        output.extend_from_slice(&xml[..position]);
        output.extend_from_slice(fragment.as_bytes());
        output.extend_from_slice(&xml[position..]);
        return Ok(output);
    }
    // No extension list: create one before the slide end tag.
    let mut output = Vec::with_capacity(xml.len() + fragment.len() + 24);
    output.extend_from_slice(&xml[..root_end]);
    output.extend_from_slice(b"<p:extLst xmlns:p=\"");
    output.extend_from_slice(slide_dialect(xml)?.as_bytes());
    output.extend_from_slice(b"\">");
    output.extend_from_slice(fragment.as_bytes());
    output.extend_from_slice(b"</p:extLst>");
    output.extend_from_slice(&xml[root_end..]);
    Ok(output)
}
