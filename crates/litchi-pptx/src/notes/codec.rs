//! `PresentationML` notes XML, text projection, and bounded validation codecs.

use super::model::{Conformance, Slide};
use super::{
    MAX_ATTRIBUTE_BYTES, MAX_ATTRIBUTES, MAX_DEPTH, MAX_NODES, MAX_NOTES_XML, allocation, invalid,
    limit, resolved, xml_error,
};
use crate::{Error, Result};
use litchi_ooxml_common::mce::process_ooxml;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

const NOTES_XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const NOTES_XML_BODY_PREFIX: &str = concat!(
    "<p:cSld><p:spTree>",
    "<p:nvGrpSpPr>",
    r#"<p:cNvPr id="1" name=""/>"#,
    "<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>",
    "<p:grpSpPr>",
    r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>"#,
    r#"<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>"#,
    "</p:grpSpPr><p:sp><p:nvSpPr>",
    r#"<p:cNvPr id="2" name="Notes Placeholder"/>"#,
    r#"<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>"#,
    r#"<p:nvPr><p:ph type="body" idx="1"/></p:nvPr>"#,
    "</p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>",
    r#"<a:rPr lang="en-US" dirty="0"/><a:t>"#,
);
const NOTES_XML_SUFFIX: &str = concat!(
    "</a:t></a:r></a:p></p:txBody></p:sp>",
    "</p:spTree></p:cSld>",
    r#"<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"#,
    "</p:notes>",
);

/// Return the deterministic Transitional notes-master producer template.
#[must_use]
pub fn master_xml() -> &'static str {
    include_str!("resources/generated/notesMaster.xml")
}

/// Encode one bounded Transitional plain-text speaker-notes slide.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn write_text(text: &str) -> Result<Vec<u8>> {
    write_text_with(Conformance::Transitional, text)
}

/// Encode one bounded plain-text speaker-notes slide in the chosen dialect.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn write_text_with(conformance: Conformance, text: &str) -> Result<Vec<u8>> {
    if text.len() > MAX_NOTES_XML {
        return Err(Error::Limit {
            resource: "speaker-notes text bytes",
            limit: MAX_NOTES_XML,
        });
    }
    if !text.chars().all(is_xml_char) {
        return Err(invalid("speaker notes contain an invalid XML character"));
    }
    let escaped = quick_xml::escape::escape(text);
    let prefix = [
        NOTES_XML_DECLARATION,
        r#"<p:notes xmlns:p=""#,
        conformance.p(),
        r#"" xmlns:a=""#,
        conformance.a(),
        r#"" xmlns:r=""#,
        conformance.r(),
        r#"">"#,
        NOTES_XML_BODY_PREFIX,
    ];
    let prefix_len = prefix
        .iter()
        .try_fold(0usize, |len, part| len.checked_add(part.len()))
        .ok_or_else(|| invalid("speaker-notes XML length overflow"))?;
    let capacity = prefix_len
        .checked_add(escaped.len())
        .and_then(|len| len.checked_add(NOTES_XML_SUFFIX.len()))
        .ok_or_else(|| invalid("speaker-notes XML length overflow"))?;
    if capacity > MAX_NOTES_XML {
        return Err(Error::Limit {
            resource: "speaker-notes XML bytes",
            limit: MAX_NOTES_XML,
        });
    }
    let mut xml = String::new();
    xml.try_reserve_exact(capacity)
        .map_err(|source| allocation("speaker-notes XML", source))?;
    for part in prefix {
        xml.push_str(part);
    }
    xml.push_str(&escaped);
    xml.push_str(NOTES_XML_SUFFIX);
    Ok(xml.into_bytes())
}

/// # Errors
///
/// Returns an error if the operation fails.
fn is_xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(value as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x1_0000..=0x10_FFFF)
}

impl Slide {
    /// Flatten the inert notes XML to its `DrawingML` text runs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn text(&self) -> Result<Option<String>> {
        let processed = process_ooxml(&self.data)?;
        let mut reader = Reader::from_reader(processed.as_ref());
        reader.config_mut().trim_text(true);
        let mut in_text = false;
        let mut value = String::new();
        loop {
            match reader.read_event() {
                Ok(Event::Start(element)) if element.local_name().as_ref() == b"t" => {
                    in_text = true;
                },
                Ok(Event::Text(text)) if in_text => {
                    let decoded = text.decode().map_err(xml_error)?;
                    let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                    if !value.is_empty() {
                        value.push('\n');
                    }
                    value.push_str(&decoded);
                },
                Ok(Event::End(element)) if element.local_name().as_ref() == b"t" => in_text = false,
                Ok(Event::Eof) => break,
                Err(error) => return Err(xml_error(error)),
                _ => {},
            }
        }
        Ok((!value.is_empty()).then_some(value))
    }
}

/// Replace the inert `DrawingML` text projection without rebuilding the notes
/// document. All markup, namespace declarations, extension branches, and
/// unrelated runs remain byte-identical; only the contents of existing
/// `a:t` elements are changed.
pub(crate) fn rewrite_text(xml: &[u8], text: &str) -> Result<Vec<u8>> {
    if text.len() > MAX_NOTES_XML {
        return Err(limit("speaker-notes text bytes", MAX_NOTES_XML));
    }
    if !text.chars().all(is_xml_char) {
        return Err(invalid("speaker notes contain an invalid XML character"));
    }
    let escaped = quick_xml::escape::escape(text);
    let spans = text_spans(xml)?;
    let Some(_first) = spans.first() else {
        return Err(invalid("notes slide has no DrawingML text run"));
    };
    let mut output = Vec::new();
    output
        .try_reserve(xml.len().saturating_add(escaped.len()))
        .map_err(|source| allocation("speaker-notes text rewrite", source))?;
    let mut cursor = 0usize;
    for (index, span) in spans.iter().enumerate() {
        output.extend_from_slice(&xml[cursor..span.start]);
        if index == 0 {
            if let Some(name) = span.empty_name.as_deref() {
                output.extend_from_slice(&xml[span.start..span.end - 1]);
                output.push(b'>');
                output.extend_from_slice(escaped.as_bytes());
                output.extend_from_slice(b"</");
                output.extend_from_slice(name);
                output.push(b'>');
            } else {
                output.extend_from_slice(escaped.as_bytes());
            }
        } else if span.empty_name.is_some() {
            output.extend_from_slice(&xml[span.start..span.end]);
        }
        cursor = span.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok(output)
}

struct TextSpan {
    start: usize,
    end: usize,
    empty_name: Option<Vec<u8>>,
}

fn text_spans(xml: &[u8]) -> Result<Vec<TextSpan>> {
    // The bounded XML scan above is authoritative for syntax and namespace
    // conformance. This byte scanner only locates replaceable text payloads,
    // avoiding a serializer pass that would rewrite opaque markup.
    let mut spans = Vec::new();
    let mut cursor = 0usize;
    while let Some(relative) = xml[cursor..].iter().position(|byte| *byte == b'<') {
        let start = cursor + relative;
        if xml
            .get(start + 1)
            .is_some_and(|byte| matches!(byte, b'!' | b'?'))
        {
            cursor = start + 2;
            continue;
        }
        if xml.get(start + 1) == Some(&b'/') {
            cursor = start + 2;
            continue;
        }
        let name_start = start + 1;
        let Some(name_end) = xml[name_start..]
            .iter()
            .position(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>'))
            .map(|offset| name_start + offset)
        else {
            return Err(invalid("unterminated notes XML element"));
        };
        let name = &xml[name_start..name_end];
        let local = name.rsplit(|byte| *byte == b':').next().unwrap_or(name);
        let end = tag_end(xml, name_end)?;
        if local != b"t" {
            cursor = end + 1;
            continue;
        }
        let self_closing = xml[name_end..=end]
            .get(..xml[name_end..=end].len().saturating_sub(1))
            .and_then(|value| value.iter().rposition(|byte| !byte.is_ascii_whitespace()))
            .is_some_and(|position| xml[name_end + position] == b'/');
        if self_closing {
            spans.push(TextSpan {
                start,
                end: end + 1,
                empty_name: Some(name.to_vec()),
            });
            cursor = end + 1;
            continue;
        }
        let close_prefix = {
            let mut value = Vec::with_capacity(name.len() + 2);
            value.extend_from_slice(b"</");
            value.extend_from_slice(name);
            value.push(b'>');
            value
        };
        let Some(close_relative) = xml[end + 1..]
            .windows(close_prefix.len())
            .position(|window| window == close_prefix)
        else {
            return Err(invalid("notes text element is unterminated"));
        };
        let close_start = end + 1 + close_relative;
        spans.push(TextSpan {
            start: end + 1,
            end: close_start,
            empty_name: None,
        });
        cursor = close_start + close_prefix.len();
    }
    Ok(spans)
}

fn tag_end(xml: &[u8], start: usize) -> Result<usize> {
    let mut quote = None;
    for (offset, byte) in xml[start..].iter().enumerate() {
        match (quote, *byte) {
            (None, b'\'' | b'"') => quote = Some(*byte),
            (Some(value), byte) if value == byte => quote = None,
            (None, b'>') => return Ok(start + offset),
            _ => {},
        }
    }
    Err(invalid("unterminated notes XML start tag"))
}

#[derive(Default)]
pub(crate) struct XmlScan {
    pub(crate) relationship_attributes: Vec<String>,
    pub(crate) notes_master_ids: Vec<String>,
    pub(crate) slide_ids: Vec<String>,
}

pub(crate) fn validate_resource_xml(
    xml: &[u8],
    max: usize,
    conformance: Conformance,
    root: &str,
    label: &str,
) -> Result<()> {
    let scan = scan_xml(xml, max, conformance, root)?;
    if !scan.relationship_attributes.is_empty() {
        return Err(invalid(format!(
            "{label} contains unsupported outbound relationship references"
        )));
    }
    Ok(())
}

pub(crate) fn root_conformance(xml: &[u8], max: usize, root: &str) -> Result<Conformance> {
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        if scan_xml(xml, max, conformance, root).is_ok() {
            return Ok(conformance);
        }
    }
    Err(invalid(format!("invalid {root} root or namespace")))
}

pub(crate) fn scan_xml(
    xml: &[u8],
    max: usize,
    conformance: Conformance,
    expected_root: &str,
) -> Result<XmlScan> {
    if xml.len() > max {
        return Err(limit("notes XML bytes", max));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > max {
        return Err(limit("processed notes XML bytes", max));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut attributes = 0usize;
    let mut attribute_bytes = 0usize;
    let mut root_seen = false;
    let mut scan = XmlScan::default();
    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                nodes += 1;
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(limit("notes XML depth", MAX_DEPTH));
                }
                if nodes > MAX_NODES {
                    return Err(limit("notes XML nodes", MAX_NODES));
                }
                inspect_element(
                    &reader,
                    &element,
                    conformance,
                    expected_root,
                    !root_seen,
                    &mut attributes,
                    &mut attribute_bytes,
                    &mut scan,
                )?;
                root_seen = true;
            },
            Event::Empty(element) => {
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(limit("notes XML nodes", MAX_NODES));
                }
                if depth >= MAX_DEPTH {
                    return Err(limit("notes XML depth", MAX_DEPTH));
                }
                inspect_element(
                    &reader,
                    &element,
                    conformance,
                    expected_root,
                    !root_seen,
                    &mut attributes,
                    &mut attribute_bytes,
                    &mut scan,
                )?;
                root_seen = true;
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("unexpected XML closing element"));
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected")),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen || depth != 0 {
        return Err(invalid("missing or unterminated XML root"));
    }
    Ok(scan)
}

#[allow(
    clippy::too_many_arguments,
    reason = "element inspector threads one slot per notes-element field"
)]
fn inspect_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    conformance: Conformance,
    expected_root: &str,
    is_root: bool,
    attributes: &mut usize,
    attribute_bytes: &mut usize,
    scan: &mut XmlScan,
) -> Result<()> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let local = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    if is_root
        && (namespace
            != if expected_root == "theme" {
                conformance.a()
            } else {
                conformance.p()
            }
            || local != expected_root)
    {
        return Err(invalid(format!(
            "invalid {expected_root} root or namespace"
        )));
    }
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = item.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        *attributes += 1;
        if *attributes > MAX_ATTRIBUTES {
            return Err(limit("notes XML attributes", MAX_ATTRIBUTES));
        }
        let (namespace, attr_local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let attr_local = std::str::from_utf8(attr_local.as_ref()).map_err(xml_error)?;
        let raw_value = std::str::from_utf8(item.value.as_ref()).map_err(xml_error)?;
        let value = quick_xml::escape::unescape(raw_value)
            .map_err(xml_error)?
            .into_owned();
        *attribute_bytes = attribute_bytes
            .checked_add(namespace.len() + attr_local.len() + value.len())
            .ok_or_else(|| invalid("notes XML attribute byte count overflow"))?;
        if *attribute_bytes > MAX_ATTRIBUTE_BYTES {
            return Err(limit("notes XML attribute bytes", MAX_ATTRIBUTE_BYTES));
        }
        if namespace == conformance.r() {
            scan.relationship_attributes.push(value.clone());
            if attr_local == "id" {
                if local == "notesMasterId" {
                    scan.notes_master_ids.push(value.clone());
                } else if local == "sldId" {
                    scan.slide_ids.push(value.clone());
                }
            }
        }
    }
    Ok(())
}
