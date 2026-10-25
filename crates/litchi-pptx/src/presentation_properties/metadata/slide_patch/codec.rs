//! Shared slide-XML patching helpers for `PowerPoint` extension storage
//! (laser traces, slide-show events, and similar `p:ext` payloads).
//!
//! All mutation is byte-level and bounded: fragments are inserted into the
//! slide's extension list (`p:extLst`) — patched into an existing list,
//! expanding an empty one, or creating one before the slide end tag — while
//! the slide's namespace dialect (transitional or Strict) is preserved.

use super::super::is_presentationml_name;
use crate::{Error, Result};
use litchi_core::xml::ReaderOrigin;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 128;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::error::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn checked_output_size(size: Option<usize>, max_output_bytes: usize) -> Result<usize> {
    let size = size.ok_or(Error::CapacityOverflow {
        resource: "updated slide XML bytes",
    })?;
    if size > max_output_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::OutputBytes,
            size,
            max_output_bytes,
        ));
    }
    if size > MAX_SLIDE_XML_BYTES {
        return Err(limit(
            crate::error::CodecLimitKind::OutputBytes,
            size,
            MAX_SLIDE_XML_BYTES,
        ));
    }
    Ok(size)
}

/// The slide's `PresentationML` namespace URI (transitional or Strict).
pub(crate) fn slide_dialect(xml: &[u8]) -> Result<&'static str> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        match reader.read_event().map_err(xml_error)? {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                return Ok(match namespace {
                    ResolveResult::Bound(Namespace(value))
                        if value == b"http://purl.oclc.org/ooxml/presentationml/main" =>
                    {
                        "http://purl.oclc.org/ooxml/presentationml/main"
                    },
                    ResolveResult::Bound(Namespace(value))
                        if value
                            == b"http://schemas.openxmlformats.org/presentationml/2006/main" =>
                    {
                        "http://schemas.openxmlformats.org/presentationml/2006/main"
                    },
                    _ => return Err(invalid("slide root has no PresentationML namespace")),
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
/// the slide end tag. The slide root must be a single `PresentationML` `sld`
/// element; DTDs and processing instructions are rejected.
pub(crate) fn insert_extension_fragment(
    xml: &[u8],
    fragment: &str,
    max_output_bytes: usize,
) -> Result<Vec<u8>> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit(
            crate::error::CodecLimitKind::InputBytes,
            xml.len(),
            MAX_SLIDE_XML_BYTES,
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut ext_lst_depth = None;
    let mut ext_lst_end = None;
    let mut empty_ext_lst: Option<(usize, usize)> = None;
    let mut root_end = None;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("slide XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        MAX_XML_NODES,
                    )
                })?;
                if nodes > MAX_XML_NODES {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        nodes,
                        MAX_XML_NODES,
                    ));
                }
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
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        MAX_XML_NODES,
                    )
                })?;
                if nodes > MAX_XML_NODES {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        nodes,
                        MAX_XML_NODES,
                    ));
                }
                if !root_seen {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid(
                            "slide XML must have one PresentationML sld root element",
                        ));
                    }
                    root_seen = true;
                }
                let empty_depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        MAX_XML_DEPTH,
                    )
                })?;
                if empty_depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        empty_depth,
                        MAX_XML_DEPTH,
                    ));
                }
                if empty_depth == 2 && is_presentationml_name(&namespace, element.name(), b"extLst")
                {
                    // Empty `<p:extLst/>`: remember its range for expansion.
                    if ext_lst_depth.replace(usize::MAX).is_some() {
                        return Err(invalid("slide has multiple extension lists"));
                    }
                    empty_ext_lst = Some((
                        start,
                        origin
                            .offset(reader.buffer_position())
                            .ok_or_else(|| invalid("slide XML offset overflow"))?,
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
        let removed_len = end
            .checked_sub(start)
            .ok_or_else(|| invalid("slide XML extension-list offsets are out of order"))?;
        let replacement_len = open
            .len()
            .checked_add(1)
            .and_then(|size| size.checked_add(fragment.len()))
            .and_then(|size| size.checked_add(b"</p:extLst>".len()));
        let size = checked_output_size(
            xml.len()
                .checked_sub(removed_len)
                .and_then(|size| size.checked_add(replacement_len?)),
            max_output_bytes,
        )?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(size)
            .map_err(|source| Error::Allocation {
                resource: "updated slide XML",
                source,
            })?;
        output.extend_from_slice(&xml[..start]);
        output.extend_from_slice(open);
        output.extend_from_slice(b">");
        output.extend_from_slice(fragment.as_bytes());
        output.extend_from_slice(b"</p:extLst>");
        output.extend_from_slice(&xml[end..]);
        if output.len() != size {
            return Err(invalid("updated slide XML length changed during emission"));
        }
        return Ok(output);
    }
    if let Some(position) = ext_lst_end {
        let size = checked_output_size(xml.len().checked_add(fragment.len()), max_output_bytes)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(size)
            .map_err(|source| Error::Allocation {
                resource: "updated slide XML",
                source,
            })?;
        output.extend_from_slice(&xml[..position]);
        output.extend_from_slice(fragment.as_bytes());
        output.extend_from_slice(&xml[position..]);
        if output.len() != size {
            return Err(invalid("updated slide XML length changed during emission"));
        }
        return Ok(output);
    }
    // No extension list: create one before the slide end tag.
    let dialect = slide_dialect(xml)?;
    let size = checked_output_size(
        xml.len()
            .checked_add(b"<p:extLst xmlns:p=\"".len())
            .and_then(|size| size.checked_add(dialect.len()))
            .and_then(|size| size.checked_add(b"\">".len()))
            .and_then(|size| size.checked_add(fragment.len()))
            .and_then(|size| size.checked_add(b"</p:extLst>".len())),
        max_output_bytes,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(size)
        .map_err(|source| Error::Allocation {
            resource: "updated slide XML",
            source,
        })?;
    output.extend_from_slice(&xml[..root_end]);
    output.extend_from_slice(b"<p:extLst xmlns:p=\"");
    output.extend_from_slice(dialect.as_bytes());
    output.extend_from_slice(b"\">");
    output.extend_from_slice(fragment.as_bytes());
    output.extend_from_slice(b"</p:extLst>");
    output.extend_from_slice(&xml[root_end..]);
    if output.len() != size {
        return Err(invalid("updated slide XML length changed during emission"));
    }
    Ok(output)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

    #[test]
    fn rejects_a_patch_that_would_exceed_the_slide_xml_budget() {
        let xml = format!(r#"<p:sld xmlns:p="{PML}"></p:sld>"#);
        let fragment = "x".repeat(MAX_SLIDE_XML_BYTES);

        let error =
            insert_extension_fragment(xml.as_bytes(), &fragment, MAX_SLIDE_XML_BYTES).unwrap_err();

        assert!(matches!(
            error,
            Error::Limit {
                kind: crate::error::LimitKind::Codec(
                    crate::error::CodecLimitKind::OutputBytes
                ),
                observed,
                maximum,
            } if observed > maximum
        ));
    }

    #[test]
    fn rejects_an_unbound_presentation_prefix() {
        let xml = b"<p:sld><p:extLst/></p:sld>";
        assert!(
            insert_extension_fragment(xml, "<p:ext uri=\"urn:test\"/>", MAX_SLIDE_XML_BYTES)
                .is_err()
        );
    }

    #[test]
    fn selected_output_budget_is_checked_before_reserve() {
        let xml = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:sld>"#;
        assert!(insert_extension_fragment(xml, "<p:ext uri=\"urn:test\"/>", xml.len()).is_err());
    }
}
