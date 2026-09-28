//! `PresentationML` notes XML, text projection, and bounded validation codecs.

use super::model::{Conformance, Slide};
use super::{
    MAX_ATTRIBUTE_BYTES, MAX_ATTRIBUTES, MAX_DEPTH, MAX_NODES, MAX_NOTES_XML, allocation, invalid,
    limit, resolved, xml_error,
};
use crate::{Error, Result};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::NamespaceResolver;
#[cfg(test)]
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
        reader.config_mut().trim_text(false);
        let mut in_text = false;
        let mut seen_text = false;
        let mut value = String::new();
        loop {
            match reader.read_event() {
                Ok(Event::Start(element)) if element.local_name().as_ref() == b"t" => {
                    if seen_text && !value.is_empty() {
                        value.push('\n');
                    }
                    seen_text = true;
                    in_text = true;
                },
                Ok(Event::Text(text)) if in_text => {
                    let decoded = text.decode().map_err(xml_error)?;
                    let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                    value.push_str(&decoded);
                },
                Ok(Event::CData(text)) if in_text => {
                    value.push_str(&text.decode().map_err(xml_error)?);
                },
                Ok(Event::GeneralRef(reference)) if in_text => {
                    value.push_str(&litchi_ooxml_common::xml::decode_xml_reference(&reference)?);
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

/// Classify one already-MCE-processed PresentationML root with the same
/// Transitional-then-Strict retry and generic-error masking as
/// [`root_conformance`].  `raw_len` is checked separately because the raw
/// notes limit is part of `scan_xml`'s contract even when the processed bytes
/// are borrowed from a capture-local reader.
pub(crate) fn root_conformance_from_processed(
    processed: &[u8],
    raw_len: usize,
    max: usize,
    root: &str,
) -> Option<Conformance> {
    [Conformance::Transitional, Conformance::Strict]
        .into_iter()
        .find(|&conformance| scan_processed_xml(processed, raw_len, max, conformance, root).is_ok())
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
    scan_processed_xml(
        processed.as_ref(),
        xml.len(),
        max,
        conformance,
        expected_root,
    )
}

fn scan_processed_xml(
    processed: &[u8],
    raw_len: usize,
    max: usize,
    conformance: Conformance,
    expected_root: &str,
) -> Result<XmlScan> {
    if raw_len > max {
        return Err(limit("notes XML bytes", max));
    }
    if processed.len() > max {
        return Err(limit("processed notes XML bytes", max));
    }
    // A slice reader's events borrow `processed` directly, so no event is
    // copied into a scratch buffer; the parser and its errors are the ones
    // the buffered read uses on the same bytes.
    let mut reader = Reader::from_reader(processed);
    reader.config_mut().trim_text(false);
    let mut resolver = NamespaceResolver::default();
    let mut pending_pop = false;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut attributes = 0usize;
    let mut attribute_bytes = 0usize;
    let mut root_seen = false;
    let mut scan = XmlScan::default();
    loop {
        if pending_pop {
            resolver.pop();
            pending_pop = false;
        }
        match reader.read_event() {
            Ok(Event::Start(ref element)) => {
                resolver
                    .push(element)
                    .map_err(quick_xml::Error::from)
                    .map_err(xml_error)?;
                nodes += 1;
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(limit("notes XML depth", MAX_DEPTH));
                }
                if nodes > MAX_NODES {
                    return Err(limit("notes XML nodes", MAX_NODES));
                }
                inspect_element(
                    &resolver,
                    element,
                    conformance,
                    expected_root,
                    !root_seen,
                    &mut attributes,
                    &mut attribute_bytes,
                    &mut scan,
                )?;
                root_seen = true;
            },
            Ok(Event::Empty(ref element)) => {
                resolver
                    .push(element)
                    .map_err(quick_xml::Error::from)
                    .map_err(xml_error)?;
                pending_pop = true;
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(limit("notes XML nodes", MAX_NODES));
                }
                if depth >= MAX_DEPTH {
                    return Err(limit("notes XML depth", MAX_DEPTH));
                }
                inspect_element(
                    &resolver,
                    element,
                    conformance,
                    expected_root,
                    !root_seen,
                    &mut attributes,
                    &mut attribute_bytes,
                    &mut scan,
                )?;
                root_seen = true;
            },
            Ok(Event::End(_)) => {
                pending_pop = true;
                if depth == 0 {
                    return Err(invalid("unexpected XML closing element"));
                }
                depth -= 1;
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::CData(_)) => return Err(invalid("CDATA is rejected")),
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(xml_error(error)),
        }
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
    resolver: &NamespaceResolver,
    element: &BytesStart<'_>,
    conformance: Conformance,
    expected_root: &str,
    is_root: bool,
    attributes: &mut usize,
    attribute_bytes: &mut usize,
    scan: &mut XmlScan,
) -> Result<()> {
    // Every name and value below is validated exactly as before but borrowed
    // from the event or the resolver; only a value the scan reports is owned.
    let namespace = resolved(resolver.resolve_element(element.name()).0)?;
    let local = std::str::from_utf8(element.local_name().into_inner()).map_err(xml_error)?;
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
    if element.attributes_raw().is_empty() {
        return Ok(());
    }
    for item in element.checked_attributes() {
        let item = item.map_err(xml_error)?;
        let raw = item.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        *attributes += 1;
        if *attributes > MAX_ATTRIBUTES {
            return Err(limit("notes XML attributes", MAX_ATTRIBUTES));
        }
        let (namespace, attr_local) = resolver.resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let attr_local = std::str::from_utf8(attr_local.into_inner()).map_err(xml_error)?;
        let raw_value = std::str::from_utf8(item.value.as_ref()).map_err(xml_error)?;
        let value = quick_xml::escape::unescape(raw_value).map_err(xml_error)?;
        *attribute_bytes = attribute_bytes
            .checked_add(namespace.len() + attr_local.len() + value.len())
            .ok_or_else(|| invalid("notes XML attribute byte count overflow"))?;
        if *attribute_bytes > MAX_ATTRIBUTE_BYTES {
            return Err(limit("notes XML attribute bytes", MAX_ATTRIBUTE_BYTES));
        }
        if namespace == conformance.r() {
            scan.relationship_attributes.push(value.as_ref().to_owned());
            if attr_local == "id" {
                if local == "notesMasterId" {
                    scan.notes_master_ids.push(value.as_ref().to_owned());
                } else if local == "sldId" {
                    scan.slide_ids.push(value.as_ref().to_owned());
                }
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    const TRANSITIONAL_ROOT: &[u8] =
        br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
    const STRICT_ROOT: &[u8] =
        br#"<p:sld xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"/>"#;
    const INVALID_ROOT: &[u8] = br#"<p:sld xmlns:p="urn:invalid"/>"#;

    #[test]
    fn processed_root_classification_matches_raw_transitional_and_strict_retry() {
        for (raw, expected) in [
            (TRANSITIONAL_ROOT, Conformance::Transitional),
            (STRICT_ROOT, Conformance::Strict),
        ] {
            let processed = process_ooxml(raw).expect("test root must process");
            assert_eq!(
                root_conformance(raw, crate::notes::MAX_SLIDE_XML, "sld")
                    .expect("raw root must classify"),
                expected
            );
            assert_eq!(
                root_conformance_from_processed(
                    processed.as_ref(),
                    raw.len(),
                    crate::notes::MAX_SLIDE_XML,
                    "sld",
                ),
                Some(expected)
            );
        }
    }

    #[test]
    fn processed_root_classification_keeps_generic_failure_masking_and_limits() {
        let processed = process_ooxml(INVALID_ROOT).expect("test root must process");
        let raw_error = root_conformance(INVALID_ROOT, crate::notes::MAX_SLIDE_XML, "sld")
            .expect_err("invalid root must be refused");
        assert_eq!(
            raw_error.to_string(),
            "invalid PresentationML: invalid sld root or namespace"
        );
        assert_eq!(
            root_conformance_from_processed(
                processed.as_ref(),
                INVALID_ROOT.len(),
                crate::notes::MAX_SLIDE_XML,
                "sld",
            ),
            None
        );

        let raw_limited = root_conformance(INVALID_ROOT, INVALID_ROOT.len() - 1, "sld")
            .expect_err("raw notes limit must be masked by root classification");
        assert_eq!(raw_limited.to_string(), raw_error.to_string());
        assert_eq!(
            root_conformance_from_processed(
                processed.as_ref(),
                INVALID_ROOT.len(),
                INVALID_ROOT.len() - 1,
                "sld",
            ),
            None
        );

        let processed_limit =
            scan_processed_xml(processed.as_ref(), 0, 1, Conformance::Transitional, "sld")
                .err()
                .expect("processed notes limit must remain visible to the scanner");
        assert!(matches!(
            processed_limit,
            Error::Limit {
                resource: "processed notes XML bytes",
                limit: 1,
            }
        ));
    }

    fn assert_empty_attribute_tail_candidate_matches_oracle(
        label: &str,
        processed: &[u8],
        conformance: Conformance,
        expected_refusal: bool,
    ) {
        let expected = outcome(buffered_scan_oracle(
            processed,
            processed.len(),
            crate::notes::MAX_SLIDE_XML,
            conformance,
            "sld",
        ));
        let actual = outcome(scan_processed_xml(
            processed,
            processed.len(),
            crate::notes::MAX_SLIDE_XML,
            conformance,
            "sld",
        ));
        assert_eq!(
            actual, expected,
            "{label}: candidate scanner diverged from the independent buffered oracle"
        );
        assert_eq!(
            expected.is_err(),
            expected_refusal,
            "{label}: boundary case did not have the expected refusal outcome"
        );
    }

    #[test]
    fn empty_attribute_tail_candidate_matches_buffered_oracle_boundaries() {
        let mut invalid_utf8_name = format!(r#"<p:sld xmlns:p="{P_NS}"><p:"#).into_bytes();
        invalid_utf8_name.extend_from_slice(b"\xff/></p:sld>");
        let cases = vec![
            (
                "Start with an exact empty attribute tail",
                format!(r#"<p:sld xmlns:p="{P_NS}"><p:x></p:x></p:sld>"#).into_bytes(),
                false,
            ),
            (
                "Empty with an exact empty attribute tail",
                format!(r#"<p:sld xmlns:p="{P_NS}"><p:x/></p:sld>"#).into_bytes(),
                false,
            ),
            (
                "space before an empty-element close",
                format!("<p:sld xmlns:p=\"{P_NS}\"><p:x /></p:sld>").into_bytes(),
                false,
            ),
            (
                "tab before an empty-element close",
                format!("<p:sld xmlns:p=\"{P_NS}\"><p:x\t/></p:sld>").into_bytes(),
                false,
            ),
            (
                "newline before an empty-element close",
                format!("<p:sld xmlns:p=\"{P_NS}\"><p:x\n/></p:sld>").into_bytes(),
                false,
            ),
            (
                "namespace declaration remains on the checked path",
                format!(r#"<p:sld xmlns:p="{P_NS}"><p:x xmlns:q="urn:vendor"/></p:sld>"#)
                    .into_bytes(),
                false,
            ),
            (
                "malformed attribute syntax",
                format!(r#"<p:sld xmlns:p="{P_NS}"><p:x a/></p:sld>"#).into_bytes(),
                true,
            ),
            (
                "malformed attribute name",
                {
                    let mut value = format!(r#"<p:sld xmlns:p="{P_NS}"><p:x "#).into_bytes();
                    value.extend_from_slice(b"\xff=\"v\"/></p:sld>");
                    value
                },
                true,
            ),
            (
                "malformed attribute value entity",
                format!(r#"<p:sld xmlns:p="{P_NS}"><p:x a="&bad;"/></p:sld>"#).into_bytes(),
                true,
            ),
            (
                "duplicate attribute",
                format!(r#"<p:sld xmlns:p="{P_NS}"><p:x a="1" a="2"/></p:sld>"#).into_bytes(),
                true,
            ),
            (
                "undeclared element prefix with no attributes",
                format!(r#"<p:sld xmlns:p="{P_NS}"><q:x/></p:sld>"#).into_bytes(),
                true,
            ),
            (
                "undeclared attribute prefix",
                format!(r#"<p:sld xmlns:p="{P_NS}"><p:x q:y="1"/></p:sld>"#).into_bytes(),
                true,
            ),
            (
                "invalid UTF-8 element name with no attributes",
                invalid_utf8_name,
                true,
            ),
            ("invalid root with no attributes", b"<sld/>".to_vec(), true),
            ("invalid root namespace", INVALID_ROOT.to_vec(), true),
        ];

        for (label, xml, expected_refusal) in cases {
            assert_empty_attribute_tail_candidate_matches_oracle(
                label,
                &xml,
                Conformance::Transitional,
                expected_refusal,
            );
        }

        let strict = format!(r#"<p:sld xmlns:p="{STRICT_NS}"><p:x/></p:sld>"#).into_bytes();
        assert_empty_attribute_tail_candidate_matches_oracle(
            "Strict empty-element path",
            &strict,
            Conformance::Strict,
            false,
        );
    }

    const STRICT_NS: &str = "http://purl.oclc.org/ooxml/presentationml/main";

    fn inspect_child_with_counters(
        xml: &[u8],
        attributes: usize,
        attribute_bytes: usize,
    ) -> (Result<()>, usize, usize) {
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        assert!(matches!(
            reader.read_event().expect("test XML root must parse"),
            Event::Start(_)
        ));
        let event = reader.read_event().expect("test XML child must parse");
        let element = match event {
            Event::Start(element) | Event::Empty(element) => element,
            _ => panic!("test XML child must be a start or empty element"),
        };
        let mut attributes_seen = attributes;
        let mut attribute_bytes_seen = attribute_bytes;
        let mut scan = XmlScan::default();
        let result = inspect_element(
            reader.resolver(),
            &element,
            Conformance::Transitional,
            "sld",
            false,
            &mut attributes_seen,
            &mut attribute_bytes_seen,
            &mut scan,
        );
        (result, attributes_seen, attribute_bytes_seen)
    }

    #[test]
    fn empty_attribute_tail_preserves_direct_attribute_limits() {
        let empty_tags = [
            format!(r#"<p:sld xmlns:p="{P_NS}"><p:x></p:x></p:sld>"#).into_bytes(),
            format!(r#"<p:sld xmlns:p="{P_NS}"><p:x/></p:sld>"#).into_bytes(),
        ];
        for xml in empty_tags {
            let (result, attributes, attribute_bytes) =
                inspect_child_with_counters(&xml, MAX_ATTRIBUTES, MAX_ATTRIBUTE_BYTES);
            assert!(result.is_ok());
            assert_eq!(attributes, MAX_ATTRIBUTES);
            assert_eq!(attribute_bytes, MAX_ATTRIBUTE_BYTES);
        }

        let attribute_tags = [
            format!(r#"<p:sld xmlns:p="{P_NS}"><p:x a="1"></p:x></p:sld>"#).into_bytes(),
            format!(r#"<p:sld xmlns:p="{P_NS}"><p:x a="1"/></p:sld>"#).into_bytes(),
        ];
        for xml in attribute_tags {
            let (result, attributes, attribute_bytes) =
                inspect_child_with_counters(&xml, MAX_ATTRIBUTES, 0);
            assert!(matches!(
                result,
                Err(Error::Limit {
                    resource: "notes XML attributes",
                    limit: MAX_ATTRIBUTES,
                })
            ));
            assert_eq!(attributes, MAX_ATTRIBUTES + 1);
            assert_eq!(attribute_bytes, 0);

            let (result, attributes, attribute_bytes) =
                inspect_child_with_counters(&xml, 0, MAX_ATTRIBUTE_BYTES);
            assert!(matches!(
                result,
                Err(Error::Limit {
                    resource: "notes XML attribute bytes",
                    limit: MAX_ATTRIBUTE_BYTES,
                })
            ));
            assert_eq!(attributes, 1);
            assert!(attribute_bytes > MAX_ATTRIBUTE_BYTES);
        }
    }

    #[test]
    fn empty_attribute_tail_preserves_node_ceiling_for_start_and_empty_children() {
        for empty in [false, true] {
            let mut xml = format!(r#"<p:sld xmlns:p="{P_NS}">"#).into_bytes();
            let child: &[u8] = if empty { b"<p:x/>" } else { b"<p:x></p:x>" };
            for _ in 0..MAX_NODES {
                xml.extend_from_slice(child);
            }
            xml.extend_from_slice(b"</p:sld>");

            let expected = outcome(buffered_scan_oracle(
                &xml,
                xml.len(),
                crate::notes::MAX_SLIDE_XML,
                Conformance::Transitional,
                "sld",
            ));
            let actual = outcome(scan_processed_xml(
                &xml,
                xml.len(),
                crate::notes::MAX_SLIDE_XML,
                Conformance::Transitional,
                "sld",
            ));
            assert_eq!(
                actual,
                expected,
                "node ceiling diverged for {} children",
                if empty { "empty" } else { "start" }
            );
            match &expected {
                Err(error) => assert!(error.contains("notes XML nodes")),
                Ok(_) => panic!("node ceiling must refuse 100001 nodes"),
            }
        }
    }

    #[test]
    fn direct_resolver_preserves_nested_empty_rebind_and_scope_pop() {
        let xml = format!(
            r#"<p:sld xmlns:p="{P_NS}" xmlns:r="{R_NS}" xmlns="urn:outer"><p:group><plain xmlns="urn:inner"><p:empty xmlns:r="urn:empty-r" xmlns="urn:empty"/></plain><p:nested xmlns:r="urn:nested-r"><p:leaf/></p:nested><p:after r:id="after"/></p:group></p:sld>"#
        )
        .into_bytes();
        let expected = outcome(buffered_scan_oracle(
            &xml,
            xml.len(),
            crate::notes::MAX_SLIDE_XML,
            Conformance::Transitional,
            "sld",
        ));
        let actual = outcome(scan_processed_xml(
            &xml,
            xml.len(),
            crate::notes::MAX_SLIDE_XML,
            Conformance::Transitional,
            "sld",
        ));
        assert_eq!(actual, expected);
        let scan = scan_processed_xml(
            &xml,
            xml.len(),
            crate::notes::MAX_SLIDE_XML,
            Conformance::Transitional,
            "sld",
        )
        .expect("nested namespace scopes must remain valid");
        assert_eq!(scan.relationship_attributes, ["after"]);
    }

    #[test]
    fn direct_resolver_reports_namespace_error_before_depth_limit() {
        let xml = format!(
            r#"<p:sld xmlns:p="{P_NS}">{nested}<p:x xmlns:xml="urn:invalid-xml"></p:x>{closes}</p:sld>"#,
            nested = "<p:x>".repeat(MAX_DEPTH - 1),
            closes = "</p:x>".repeat(MAX_DEPTH),
        )
        .into_bytes();
        let expected = outcome(buffered_scan_oracle(
            &xml,
            xml.len(),
            crate::notes::MAX_SLIDE_XML,
            Conformance::Transitional,
            "sld",
        ));
        let actual = outcome(scan_processed_xml(
            &xml,
            xml.len(),
            crate::notes::MAX_SLIDE_XML,
            Conformance::Transitional,
            "sld",
        ));
        assert_eq!(actual, expected);
        assert!(
            actual
                .expect_err("reserved namespace must be rejected")
                .contains("namespace prefix 'xml'")
        );
    }

    #[test]
    fn direct_resolver_preserves_reserved_namespace_and_declaration_cap_order() {
        for (xml, needle) in [
            (
                format!(r#"<p:sld xmlns:p="{P_NS}" xmlns:xml="urn:invalid-xml"/>"#),
                "namespace prefix 'xml'",
            ),
            (
                format!(r#"<p:sld xmlns:p="{P_NS}" xmlns:xmlns="urn:invalid-xmlns"/>"#),
                "namespace prefix 'xmlns'",
            ),
        ] {
            let xml = xml.into_bytes();
            let expected = outcome(buffered_scan_oracle(
                &xml,
                xml.len(),
                crate::notes::MAX_SLIDE_XML,
                Conformance::Transitional,
                "sld",
            ));
            let actual = outcome(scan_processed_xml(
                &xml,
                xml.len(),
                crate::notes::MAX_SLIDE_XML,
                Conformance::Transitional,
                "sld",
            ));
            assert_eq!(actual, expected);
            assert!(
                actual
                    .expect_err("reserved namespace must be rejected")
                    .contains(needle)
            );
        }

        let mut declarations = format!(r#"xmlns:p="{P_NS}""#);
        for index in 0..255 {
            declarations.push_str(&format!(r##" xmlns:n{index}="urn:n{index}""##));
        }
        declarations.push_str(r##" xmlns:xml="urn:invalid-xml""##);
        let xml = format!(r#"<p:sld {declarations}/>"#).into_bytes();
        let expected = outcome(buffered_scan_oracle(
            &xml,
            xml.len(),
            crate::notes::MAX_SLIDE_XML,
            Conformance::Transitional,
            "sld",
        ));
        let actual = outcome(scan_processed_xml(
            &xml,
            xml.len(),
            crate::notes::MAX_SLIDE_XML,
            Conformance::Transitional,
            "sld",
        ));
        assert_eq!(actual, expected);
        assert!(
            actual
                .expect_err("the declaration cap must be reached before the reserved binding")
                .contains("more than 256 namespace bindings")
        );
    }

    /// The scanner as it stood before change 0743, verbatim: buffered reads
    /// into a scratch vector and owned names and values. The borrowed scanner
    /// must return exactly its values and exactly its refusals.
    fn buffered_scan_oracle(
        processed: &[u8],
        raw_len: usize,
        max: usize,
        conformance: Conformance,
        expected_root: &str,
    ) -> Result<XmlScan> {
        if raw_len > max {
            return Err(limit("notes XML bytes", max));
        }
        if processed.len() > max {
            return Err(limit("processed notes XML bytes", max));
        }
        let mut reader = NsReader::from_reader(processed);
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
                    inspect_element_oracle(
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
                    inspect_element_oracle(
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

    fn resolved_owned_oracle(value: quick_xml::name::ResolveResult<'_>) -> Result<String> {
        match value {
            quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(value)) => {
                Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
            },
            quick_xml::name::ResolveResult::Unbound => Ok(String::new()),
            quick_xml::name::ResolveResult::Unknown(prefix) => Err(invalid(format!(
                "unbound XML prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            ))),
        }
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "verbatim oracle of the pre-change element inspector"
    )]
    fn inspect_element_oracle(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        conformance: Conformance,
        expected_root: &str,
        is_root: bool,
        attributes: &mut usize,
        attribute_bytes: &mut usize,
        scan: &mut XmlScan,
    ) -> Result<()> {
        let namespace = resolved_owned_oracle(reader.resolver().resolve_element(element.name()).0)?;
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
            let namespace = resolved_owned_oracle(namespace)?;
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

    type ScanOutcome = std::result::Result<(Vec<String>, Vec<String>, Vec<String>), String>;

    fn outcome(result: Result<XmlScan>) -> ScanOutcome {
        result
            .map(|scan| {
                (
                    scan.relationship_attributes,
                    scan.notes_master_ids,
                    scan.slide_ids,
                )
            })
            .map_err(|error| format!("{error:?}"))
    }

    /// Compare the borrowed scanner with the buffered oracle for one input
    /// under every root, conformance and a spread of byte ceilings. Returns
    /// the number of comparisons and of those the oracle refused.
    fn assert_scanners_agree(label: &str, xml: &[u8]) -> (usize, usize) {
        let mut compared = 0;
        let mut refused = 0;
        let ceilings = [
            crate::notes::MAX_SLIDE_XML,
            xml.len(),
            xml.len().saturating_sub(1),
            xml.len() / 2,
            0,
        ];
        let actual_root = document_root_local_name(xml);
        let mut roots = vec!["sld", "presentation", "notes", "notesMaster", "theme"];
        if let Some(actual) = actual_root.as_deref()
            && !roots.contains(&actual)
        {
            roots.push(actual);
        }
        for root in roots {
            for conformance in [Conformance::Transitional, Conformance::Strict] {
                for (index, max) in ceilings.into_iter().enumerate() {
                    // The raw and processed lengths are checked separately;
                    // vary them independently on the first two ceilings.
                    for raw_len in [xml.len(), max.saturating_add(usize::from(index == 1))] {
                        let expected =
                            outcome(buffered_scan_oracle(xml, raw_len, max, conformance, root));
                        let actual =
                            outcome(scan_processed_xml(xml, raw_len, max, conformance, root));
                        assert_eq!(
                            actual, expected,
                            "{label}: root {root}, {conformance:?}, max {max}, raw {raw_len}"
                        );
                        compared += 1;
                        refused += usize::from(expected.is_err());
                    }
                }
            }
        }
        (compared, refused)
    }

    /// Local name of the first element, so every well-formed input is also
    /// scanned under the root it actually has.
    fn document_root_local_name(xml: &[u8]) -> Option<String> {
        let mut reader = Reader::from_reader(xml);
        loop {
            match reader.read_event().ok()? {
                Event::Start(element) | Event::Empty(element) => {
                    return String::from_utf8(element.local_name().as_ref().to_vec()).ok();
                },
                Event::Eof => return None,
                _ => {},
            }
        }
    }

    const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn handcrafted_seeds() -> Vec<Vec<u8>> {
        let deep = format!(
            r#"<p:sld xmlns:p="{P_NS}">{}{}</p:sld>"#,
            "<p:x>".repeat(MAX_DEPTH + 1),
            "</p:x>".repeat(MAX_DEPTH + 1)
        );
        let deep_empty = format!(
            r#"<p:sld xmlns:p="{P_NS}">{}<p:y/>{}</p:sld>"#,
            "<p:x>".repeat(MAX_DEPTH - 1),
            "</p:x>".repeat(MAX_DEPTH - 1)
        );
        let mut seeds: Vec<Vec<u8>> = [
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:notesMasterIdLst><p:notesMasterId r:id="rId9"/></p:notesMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="r&amp;Id&#x33;"/></p:sldIdLst></p:presentation>"#
            ),
            format!(r#"<p:sld xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:cSld name="a &lt; b"><p:spTree/></p:cSld><p:pic r:embed="rId5" r:link="&#65;"/></p:sld>"#),
            format!(r#"<sld xmlns="{P_NS}"><cSld/></sld>"#),
            format!(r#"<p:sld xmlns:p="{P_NS}"><p:cSld xmlns:p=""><p:x/></p:cSld></p:sld>"#),
            format!(r#"<p:sld xmlns:p="{P_NS}"><q:x/></p:sld>"#),
            format!(r#"<p:sld xmlns:p="{P_NS}"><p:x q:y="1"/></p:sld>"#),
            format!(r#"<p:sld xmlns:p="{P_NS}"><p:x a="1" a="2"/></p:sld>"#),
            format!(r#"<p:sld xmlns:p="{P_NS}"><p:x a="&bad;"/></p:sld>"#),
            format!(r#"<p:sld xmlns:p="{P_NS}"><p:x a="&#0;"/></p:sld>"#),
            format!(r#"<p:sld xmlns:p="{P_NS}"><![CDATA[x]]></p:sld>"#),
            format!(r#"<p:sld xmlns:p="{P_NS}"><?pi x?></p:sld>"#),
            format!(r#"<!DOCTYPE x><p:sld xmlns:p="{P_NS}"/>"#),
            format!(r#"<p:sld xmlns:p="{P_NS}"><!-- note --> text &amp; more</p:sld> "#),
            format!(r#"<p:sld xmlns:p="{P_NS}"/><p:sld xmlns:p="{P_NS}"/>"#),
            format!(r#"<p:sld xmlns:p="{P_NS}"><p:a></p:b></p:sld>"#),
            format!(r#"<p:sld xmlns:p="{P_NS}"></p:sld></p:sld>"#),
            format!(r#"<p:sld xmlns:p="{P_NS}"><p:a>"#),
            format!("\u{feff}<p:sld xmlns:p=\"{P_NS}\"/>"),
            deep,
            deep_empty,
            String::new(),
            "   ".to_owned(),
        ]
        .into_iter()
        .map(String::into_bytes)
        .collect();
        let mut invalid_utf8 = format!(r#"<p:sld xmlns:p="{P_NS}"><p:x v="#).into_bytes();
        invalid_utf8.extend_from_slice(b"\"\xff\xfe\"/></p:sld>");
        seeds.push(invalid_utf8);
        let mut invalid_name = format!(r#"<p:sld xmlns:p="{P_NS}"><p:"#).into_bytes();
        invalid_name.extend_from_slice(b"\xc3\x28/></p:sld>");
        seeds.push(invalid_name);
        seeds
    }

    /// Deterministic structural mutations: truncations, byte substitutions
    /// with markup-significant or invalid bytes, and snippet insertions.
    fn mutations(seed: &[u8], budget: usize) -> Vec<Vec<u8>> {
        const BYTES: &[u8] = b"<>&\"'/!?=:\x00\xff\xc3 x";
        const SNIPPETS: &[&[u8]] = &[
            b"<![CDATA[c]]>",
            b"<?pi x?>",
            b"<!DOCTYPE d>",
            b"<!-- c -->",
            b"<q:x/>",
            b"<p:y r:id=\"rId7\"/>",
            b"&amp;",
            b"&bad;",
            b"</p:z>",
            b" a=\"1\"",
        ];
        let mut output = Vec::new();
        if seed.is_empty() {
            return output;
        }
        let step = (seed.len() / budget.max(1)).max(1);
        let mut state = 0x9e37_79b9_7f4a_7c15_u64 ^ seed.len() as u64;
        let mut next = || {
            state = state
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            usize::try_from(state >> 33).unwrap_or(0)
        };
        for position in (0..seed.len()).step_by(step) {
            output.push(seed[..position].to_vec());
            let mut replaced = seed.to_vec();
            replaced[position] = BYTES[next() % BYTES.len()];
            output.push(replaced);
            let mut inserted = seed[..position].to_vec();
            inserted.extend_from_slice(SNIPPETS[next() % SNIPPETS.len()]);
            inserted.extend_from_slice(&seed[position..]);
            output.push(inserted);
        }
        output
    }

    fn corpus_xml_parts(limit_parts: usize) -> Vec<Vec<u8>> {
        fn collect(directory: &std::path::Path, found: &mut Vec<std::path::PathBuf>) {
            let Ok(entries) = std::fs::read_dir(directory) else {
                return;
            };
            for entry in entries.flatten() {
                let path = entry.path();
                if path.is_dir() {
                    collect(&path, found);
                } else if path
                    .extension()
                    .is_some_and(|extension| extension == "pptx")
                {
                    found.push(path);
                }
            }
        }
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
        let mut fixtures = Vec::new();
        collect(&root, &mut fixtures);
        fixtures.sort();
        let mut parts = Vec::new();
        for fixture in fixtures {
            let Ok(bytes) = std::fs::read(&fixture) else {
                continue;
            };
            let Ok(package) = litchi_opc::OpcPackage::from_vec(bytes) else {
                continue;
            };
            // Package parts iterate in hash order; take them by name so the
            // oracle compares the same parts on every run.
            let mut named: Vec<_> = package.try_iter_parts().flatten().collect();
            named.sort_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));
            for part in named {
                if part.content_type().ends_with("+xml") && parts.len() < limit_parts {
                    parts.push(part.blob().to_vec());
                }
            }
        }
        parts
    }

    #[test]
    fn the_borrowed_scanner_matches_the_buffered_oracle_on_handcrafted_and_mutated_xml() {
        let mut compared = 0;
        let mut refused = 0;
        for (index, seed) in handcrafted_seeds().iter().enumerate() {
            let (count, errors) = assert_scanners_agree(&format!("seed {index}"), seed);
            compared += count;
            refused += errors;
            for (variant, mutated) in mutations(seed, 40).iter().enumerate() {
                let (count, errors) =
                    assert_scanners_agree(&format!("seed {index} variant {variant}"), mutated);
                compared += count;
                refused += errors;
            }
        }
        println!(
            "0743-notes-scan-oracle handcrafted compared={compared} accepted={}",
            compared - refused
        );
        assert!(refused > 0 && refused < compared);
    }

    #[test]
    fn the_borrowed_scanner_matches_the_buffered_oracle_on_the_pptx_corpus() {
        let parts = corpus_xml_parts(600);
        assert!(parts.len() >= 300, "expected the repository PPTX XML parts");
        let mut compared = 0;
        let mut refused = 0;
        for (index, part) in parts.iter().enumerate() {
            let (count, errors) = assert_scanners_agree(&format!("part {index}"), part);
            compared += count;
            refused += errors;
            if index % 10 == 0 {
                for (variant, mutated) in mutations(part, 12).iter().enumerate() {
                    let (count, errors) =
                        assert_scanners_agree(&format!("part {index} variant {variant}"), mutated);
                    compared += count;
                    refused += errors;
                }
            }
        }
        println!(
            "0743-notes-scan-oracle corpus parts={} compared={compared} accepted={}",
            parts.len(),
            compared - refused
        );
        assert!(refused > 0 && refused < compared);
    }
}
