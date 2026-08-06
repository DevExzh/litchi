//! Bounded, loss-preserving XML codec for `w12:collapsed` in a paragraph.

use crate::error::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt::Write;

use super::model::Collapsed;
use super::validation::{
    MAX_XML_BYTES, MAX_XML_DEPTH, MAX_XML_NODES, WORD_2012_NAMESPACE, parse_on_off, validate,
};
use crate::paragraph::codec::is_fragment_word_name;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ByteRange {
    start: usize,
    end: usize,
}

#[derive(Debug, Default)]
struct Layout {
    root_prefix: Vec<u8>,
    root_start_end: Option<usize>,
    root_empty: Option<ByteRange>,
    ppr_empty: Option<ByteRange>,
    ppr_close_start: Option<usize>,
    collapsed: Option<ByteRange>,
}

/// Read the direct Word 2012 collapse marker from a paragraph.
pub(crate) fn read(xml_bytes: &[u8]) -> Result<Option<Collapsed>> {
    let (_, value) = locate(xml_bytes)?;
    Ok(value)
}

/// Replace the direct collapse marker while preserving every unrelated byte.
pub(crate) fn rewrite(xml_bytes: &[u8], value: Option<Collapsed>) -> Result<Vec<u8>> {
    validate(value)?;
    let (layout, current) = locate(xml_bytes)?;
    if current == value {
        return Ok(xml_bytes.to_vec());
    }

    let replacement = value.map(render_collapsed).transpose()?;
    match (layout.collapsed, replacement) {
        (Some(range), Some(replacement)) => Ok(splice(xml_bytes, range, &replacement)),
        (Some(range), None) => Ok(splice(xml_bytes, range, &[])),
        (None, None) => Ok(xml_bytes.to_vec()),
        (None, Some(replacement)) => {
            if let Some(close_start) = layout.ppr_close_start {
                return Ok(insert_at(xml_bytes, close_start, &replacement));
            }
            if let Some(range) = layout.ppr_empty {
                let ppr_name = qualified_name(&layout.root_prefix, "pPr")?;
                let close = format!("</{ppr_name}>");
                return expand_empty(xml_bytes, range, &replacement, close.as_bytes());
            }

            let ppr_name = qualified_name(&layout.root_prefix, "pPr")?;
            let ppr = format!("<{ppr_name}>{}</{ppr_name}>", text(&replacement));
            if let Some(root_start_end) = layout.root_start_end {
                return Ok(insert_at(xml_bytes, root_start_end, ppr.as_bytes()));
            }
            let root_range = layout.root_empty.ok_or_else(|| {
                Error::InvalidFormat("paragraph XML has no editable root".to_owned())
            })?;
            let root_name = qualified_name(&layout.root_prefix, "p")?;
            let close = format!("</{root_name}>");
            expand_empty(xml_bytes, root_range, ppr.as_bytes(), close.as_bytes())
        },
    }
}

/// Append a canonical collapse marker to a generated `w:pPr`.
pub(crate) fn append_xml(xml: &mut String, value: Collapsed) -> Result<()> {
    write!(
        xml,
        r#"<w12:collapsed xmlns:w12="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w12" w12:val="{}"/>"#,
        String::from_utf8_lossy(WORD_2012_NAMESPACE),
        value.as_xml()
    )?;
    Ok(())
}

fn locate(xml_bytes: &[u8]) -> Result<(Layout, Option<Collapsed>)> {
    if xml_bytes.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "paragraph XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }

    let mut reader = NsReader::from_reader(xml_bytes);
    reader.config_mut().trim_text(false);
    let mut layout = Layout::default();
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut ppr_depth = None;
    let mut collapsed_depth = None;
    let mut collapsed_start = None;
    let mut value = None;
    let mut nodes = 0usize;

    loop {
        let event_start = offset(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event = event.into_owned();
        let event_end = offset(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("paragraph XML element counter overflow".to_owned())
            })?;
            if nodes > MAX_XML_NODES {
                return Err(Error::InvalidFormat(format!(
                    "paragraph XML exceeds {MAX_XML_NODES} elements"
                )));
            }
        }

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("paragraph XML nesting is too deep".to_owned())
                })?;
                if depth > MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "paragraph XML nesting exceeds {MAX_XML_DEPTH}"
                    )));
                }
                if depth == 1 {
                    validate_root(
                        &mut layout,
                        &mut fragment_prefix,
                        &mut saw_root,
                        &namespace,
                        &element,
                        event_end,
                        None,
                    )?;
                } else {
                    let is_word = is_fragment_word_name(
                        &namespace,
                        element.name(),
                        element.local_name().as_ref(),
                        &fragment_prefix,
                    );
                    if depth == 2 && is_word && element.local_name().as_ref() == b"pPr" {
                        if ppr_depth.is_some() || layout.ppr_empty.is_some() {
                            return Err(Error::InvalidFormat(
                                "paragraph has duplicate pPr".to_owned(),
                            ));
                        }
                        ppr_depth = Some(depth);
                    } else if depth == 3
                        && ppr_depth == Some(2)
                        && is_word2012(&namespace)
                        && element.local_name().as_ref() == b"collapsed"
                    {
                        if layout.collapsed.is_some() || collapsed_depth.is_some() {
                            return Err(Error::InvalidFormat(
                                "paragraph has duplicate Word 2012 collapsed elements".to_owned(),
                            ));
                        }
                        value = Some(parse_collapsed(&element, reader.decoder(), &resolver)?);
                        collapsed_depth = Some(depth);
                        collapsed_start = Some(event_start);
                    }
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("paragraph XML nesting is too deep".to_owned())
                })?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "paragraph XML nesting exceeds {MAX_XML_DEPTH}"
                    )));
                }
                if child_depth == 1 {
                    validate_root(
                        &mut layout,
                        &mut fragment_prefix,
                        &mut saw_root,
                        &namespace,
                        &element,
                        event_end,
                        Some(ByteRange {
                            start: event_start,
                            end: event_end,
                        }),
                    )?;
                } else {
                    let is_word = is_fragment_word_name(
                        &namespace,
                        element.name(),
                        element.local_name().as_ref(),
                        &fragment_prefix,
                    );
                    if child_depth == 2 && is_word && element.local_name().as_ref() == b"pPr" {
                        if ppr_depth.is_some() || layout.ppr_empty.is_some() {
                            return Err(Error::InvalidFormat(
                                "paragraph has duplicate pPr".to_owned(),
                            ));
                        }
                        layout.ppr_empty = Some(ByteRange {
                            start: event_start,
                            end: event_end,
                        });
                    } else if child_depth == 3
                        && ppr_depth == Some(2)
                        && is_word2012(&namespace)
                        && element.local_name().as_ref() == b"collapsed"
                    {
                        if layout.collapsed.is_some() {
                            return Err(Error::InvalidFormat(
                                "paragraph has duplicate Word 2012 collapsed elements".to_owned(),
                            ));
                        }
                        layout.collapsed = Some(ByteRange {
                            start: event_start,
                            end: event_end,
                        });
                        value = Some(parse_collapsed(&element, reader.decoder(), &resolver)?);
                    }
                }
            },
            Event::End(element) => {
                if collapsed_depth == Some(depth) {
                    let start = collapsed_start.take().ok_or_else(|| {
                        Error::InvalidFormat("collapsed element has no start".to_owned())
                    })?;
                    layout.collapsed = Some(ByteRange {
                        start,
                        end: event_end,
                    });
                    collapsed_depth = None;
                }
                if ppr_depth == Some(depth) {
                    ppr_depth = None;
                    layout.ppr_close_start = Some(event_start);
                }
                let _ = element;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid paragraph XML nesting".to_owned())
                })?;
            },
            Event::Eof if depth != 0 || ppr_depth.is_some() || collapsed_depth.is_some() => {
                return Err(Error::InvalidFormat(
                    "unterminated paragraph XML".to_owned(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !saw_root {
        return Err(Error::InvalidFormat("paragraph XML has no root".to_owned()));
    }
    Ok((layout, value))
}

fn validate_root(
    layout: &mut Layout,
    fragment_prefix: &mut Option<Option<Vec<u8>>>,
    saw_root: &mut bool,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    event_end: usize,
    empty: Option<ByteRange>,
) -> Result<()> {
    if *saw_root {
        return Err(Error::InvalidFormat(
            "paragraph XML has multiple roots".to_owned(),
        ));
    }
    if !matches!(namespace, ResolveResult::Bound(_)) {
        *fragment_prefix = Some(
            element
                .name()
                .prefix()
                .map(|prefix| prefix.into_inner().to_vec()),
        );
    }
    if element.local_name().as_ref() != b"p"
        || !is_fragment_word_name(namespace, element.name(), b"p", fragment_prefix)
    {
        return Err(Error::InvalidFormat(
            "paragraph XML has an invalid root".to_owned(),
        ));
    }
    *saw_root = true;
    layout.root_prefix = element_prefix(element);
    if let Some(empty) = empty {
        layout.root_empty = Some(empty);
    } else {
        layout.root_start_end = Some(event_end);
    }
    Ok(())
}

fn parse_collapsed(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Collapsed> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_word2012_attribute(&namespace, element) {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(
                "collapsed has duplicate val attributes".to_owned(),
            ));
        }
        let decoded = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        value = Some(decoded.into_owned());
    }
    parse_on_off(value.as_deref())
}

fn is_word2012(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == WORD_2012_NAMESPACE
    )
}

fn is_word2012_attribute(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    if is_word2012(namespace) || matches!(namespace, ResolveResult::Unbound) {
        return true;
    }
    matches!(
        namespace,
        ResolveResult::Unknown(prefix)
            if element
                .name()
                .prefix()
                .is_some_and(|element_prefix| element_prefix.as_ref() == prefix.as_slice())
    )
}

fn element_prefix(element: &BytesStart<'_>) -> Vec<u8> {
    element
        .name()
        .prefix()
        .map_or_else(Vec::new, |prefix| prefix.into_inner().to_vec())
}

fn qualified_name(prefix: &[u8], local: &str) -> Result<String> {
    let prefix = std::str::from_utf8(prefix)
        .map_err(|error| Error::InvalidFormat(format!("invalid paragraph prefix: {error}")))?;
    if prefix.is_empty() {
        Ok(local.to_owned())
    } else {
        Ok(format!("{prefix}:{local}"))
    }
}

fn render_collapsed(value: Collapsed) -> Result<Vec<u8>> {
    let mut xml = String::new();
    append_xml(&mut xml, value)?;
    Ok(xml.into_bytes())
}

fn text(bytes: &[u8]) -> &str {
    // `render_collapsed` only emits ASCII, so this conversion cannot fail.
    std::str::from_utf8(bytes).expect("canonical collapsed XML is UTF-8")
}

fn offset(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| Error::InvalidFormat("paragraph XML offset does not fit usize".to_owned()))
}

fn splice(source: &[u8], range: ByteRange, replacement: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(
        source
            .len()
            .saturating_sub(range.end.saturating_sub(range.start))
            .saturating_add(replacement.len()),
    );
    output.extend_from_slice(&source[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[range.end..]);
    output
}

fn insert_at(source: &[u8], offset: usize, insertion: &[u8]) -> Vec<u8> {
    splice(
        source,
        ByteRange {
            start: offset,
            end: offset,
        },
        insertion,
    )
}

fn expand_empty(source: &[u8], range: ByteRange, inner: &[u8], closing: &[u8]) -> Result<Vec<u8>> {
    let raw = &source[range.start..range.end];
    let close = raw
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| Error::InvalidFormat("empty paragraph element has no close".to_owned()))?;
    let slash = raw[..close]
        .iter()
        .rposition(|byte| !byte.is_ascii_whitespace())
        .filter(|&index| raw[index] == b'/')
        .ok_or_else(|| {
            Error::InvalidFormat("paragraph empty element is missing '/>'".to_owned())
        })?;

    let mut replacement = Vec::with_capacity(raw.len() + inner.len() + closing.len());
    replacement.extend_from_slice(&raw[..slash]);
    replacement.push(b'>');
    replacement.extend_from_slice(inner);
    replacement.extend_from_slice(closing);
    Ok(splice(source, range, &replacement))
}
