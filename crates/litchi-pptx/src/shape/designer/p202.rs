//! Strict payload codec for `PowerPoint` 2020 Designer metadata.

use std::ops::Range;

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::{DrawingProperties, Limits, P202_NAMESPACE, Tag, Tags};
use crate::{Error, Result};

const P202: &[u8] = P202_NAMESPACE.as_bytes();

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct PropertiesSource {
    pub(crate) value: DrawingProperties,
    pub(crate) inner_extensions: Option<Vec<u8>>,
}

pub(crate) fn read_properties(xml: &[u8], limits: Limits) -> Result<PropertiesSource> {
    read_properties_with_prefix(xml, limits, None)
}

/// Parse a detached range whose unresolved prefixes were proven by a
/// full-owner namespace-aware scanner.
///
/// `proven_prefixes` is a NUL-separated set of prefixes bound to the p202
/// namespace in the complete owner scope. An empty slice proves the default
/// namespace. Passing a prefix that was not established by the owner is never
/// permitted.
pub(crate) fn read_properties_with_prefix(
    xml: &[u8],
    limits: Limits,
    proven_prefix: Option<&[u8]>,
) -> Result<PropertiesSource> {
    parse(xml, limits, proven_prefix, Root::Properties).map(|parsed| PropertiesSource {
        value: DrawingProperties::new()
            .with_editable(parsed.editable)
            .with_tags(parsed.tags, limits)
            .expect("parsed tags satisfy the same limits"),
        inner_extensions: parsed.inner_extensions,
    })
}

pub(crate) fn read_tags(xml: &[u8], limits: Limits) -> Result<Tags> {
    read_tags_with_prefix(xml, limits, None)
}

pub(crate) fn read_tags_with_prefix(
    xml: &[u8],
    limits: Limits,
    proven_prefix: Option<&[u8]>,
) -> Result<Tags> {
    Ok(parse(xml, limits, proven_prefix, Root::Tags)?
        .tags
        .expect("tag-list parser always produces a present list"))
}

pub(crate) fn write_properties(
    value: &DrawingProperties,
    inner_extensions: Option<&[u8]>,
    limits: Limits,
) -> Result<Vec<u8>> {
    let mut output = Output::new(limits);
    output.push(b"<p202:designPr xmlns:p202=\"")?;
    output.push(P202)?;
    output.push(b"\"")?;
    if let Some(editable) = value.editable() {
        output.push(if editable {
            b" edtDesignElem=\"true\""
        } else {
            b" edtDesignElem=\"false\""
        })?;
    }
    if value.tags().is_none() && inner_extensions.is_none() {
        output.push(b"/>")?;
        return Ok(output.finish());
    }
    output.push(b">")?;
    if let Some(tags) = value.tags() {
        write_tags_into(&mut output, tags, false)?;
    }
    if let Some(extension) = inner_extensions {
        if extension.len() > limits.xml_bytes() {
            return Err(limit("designer inner extension bytes", limits.xml_bytes()));
        }
        output.push(extension)?;
    }
    output.push(b"</p202:designPr>")?;
    Ok(output.finish())
}

pub(crate) fn write_tags(value: &Tags, limits: Limits) -> Result<Vec<u8>> {
    value.validate(limits)?;
    let mut output = Output::new(limits);
    write_tags_into(&mut output, value, true)?;
    Ok(output.finish())
}

fn write_tags_into(output: &mut Output, value: &Tags, bind_namespace: bool) -> Result<()> {
    value.validate(output.limits)?;
    output.push(if bind_namespace {
        b"<p202:designTagLst xmlns:p202=\""
    } else {
        b"<p202:designTagLst"
    })?;
    if bind_namespace {
        output.push(P202)?;
        output.push(b"\"")?;
    }
    if value.is_empty() {
        output.push(b"/>")?;
        return Ok(());
    }
    output.push(b">")?;
    for tag in value {
        output.push(b"<p202:designTag name=\"")?;
        output.push_escaped(tag.name())?;
        output.push(b"\" val=\"")?;
        output.push_escaped(tag.value())?;
        output.push(b"\"/>")?;
    }
    output.push(b"</p202:designTagLst>")
}

impl<'a> IntoIterator for &'a Tags {
    type Item = &'a Tag;
    type IntoIter = std::slice::Iter<'a, Tag>;

    fn into_iter(self) -> Self::IntoIter {
        self.as_slice().iter()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Root {
    Properties,
    Tags,
}

#[derive(Debug)]
struct Parsed {
    editable: Option<bool>,
    tags: Option<Tags>,
    inner_extensions: Option<Vec<u8>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Frame {
    Properties,
    Tags,
    Tag,
    InnerExtensions,
    Opaque,
}

fn parse(xml: &[u8], limits: Limits, proven_prefix: Option<&[u8]>, root: Root) -> Result<Parsed> {
    if xml.is_empty() {
        return Err(invalid("Designer payload is empty"));
    }
    if xml.len() > limits.xml_bytes() {
        return Err(limit("Designer payload bytes", limits.xml_bytes()));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut frames = Vec::new();
    frames
        .try_reserve_exact(8)
        .map_err(|source| Error::Allocation {
            resource: "Designer XML frames",
            source,
        })?;
    let mut events = 0usize;
    let mut parsed = Parsed {
        editable: None,
        tags: None,
        inner_extensions: None,
    };
    let mut inner_start = None;
    let mut root_closed = false;

    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let event = event.into_owned();
        if !matches!(event, Event::Eof) {
            events = events
                .checked_add(1)
                .ok_or_else(|| limit("Designer XML events", limits.xml_nodes()))?;
            if events > limits.xml_nodes() {
                return Err(limit("Designer XML events", limits.xml_nodes()));
            }
        }
        let namespace_ok = is_p202(namespace, proven_prefix);
        let empty = matches!(&event, Event::Empty(_));
        let end = position(&reader)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let local = element.local_name();
                let frame = if frames.is_empty() {
                    let expected = match root {
                        Root::Properties => b"designPr".as_slice(),
                        Root::Tags => b"designTagLst".as_slice(),
                    };
                    if !namespace_ok || local.as_ref() != expected {
                        return Err(invalid(
                            "Designer payload has the wrong root namespace or name",
                        ));
                    }
                    match root {
                        Root::Properties => {
                            parsed.editable = parse_editable(&element, decoder, limits)?;
                            Frame::Properties
                        },
                        Root::Tags => {
                            validate_no_attributes(&element, limits)?;
                            parsed.tags = Some(Tags::new());
                            Frame::Tags
                        },
                    }
                } else if frames
                    .iter()
                    .any(|frame| matches!(frame, Frame::InnerExtensions | Frame::Opaque))
                {
                    Frame::Opaque
                } else {
                    match (frames.last().copied(), namespace_ok, local.as_ref()) {
                        (Some(Frame::Properties), true, b"designTagLst") => {
                            if parsed.tags.is_some() || parsed.inner_extensions.is_some() {
                                return Err(invalid("designTagLst is duplicate or out of order"));
                            }
                            validate_no_attributes(&element, limits)?;
                            parsed.tags = Some(Tags::new());
                            Frame::Tags
                        },
                        (Some(Frame::Properties), true, b"extLst") => {
                            if parsed.inner_extensions.is_some() {
                                return Err(invalid(
                                    "designPr has duplicate inner extLst elements",
                                ));
                            }
                            validate_no_attributes(&element, limits)?;
                            inner_start = Some(start);
                            Frame::InnerExtensions
                        },
                        (Some(Frame::Tags), true, b"designTag") => {
                            let tag = parse_tag(&element, decoder, limits)?;
                            parsed
                                .tags
                                .as_mut()
                                .expect("tag parent creates a collection")
                                .push_with_limits(tag, limits)?;
                            Frame::Tag
                        },
                        _ => {
                            return Err(invalid("Designer payload has an unexpected child"));
                        },
                    }
                };
                if frames.len() >= limits.xml_depth() {
                    return Err(limit("Designer XML depth", limits.xml_depth()));
                }
                if empty {
                    finish_frame(xml, frame, start..end, &mut parsed, &mut inner_start)?;
                    if frames.is_empty() {
                        root_closed = true;
                    }
                } else {
                    frames.push(frame);
                }
            },
            Event::End(_) => {
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid("Designer XML stack underflow"))?;
                let span = inner_start.unwrap_or(start)..end;
                finish_frame(xml, frame, span, &mut parsed, &mut inner_start)?;
                if frames.is_empty() {
                    root_closed = true;
                    if end != xml.len() {
                        return Err(invalid("Designer payload has trailing markup"));
                    }
                    break;
                }
            },
            Event::Text(text) => {
                let opaque = frames.contains(&Frame::Opaque);
                if !opaque && !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("Designer payload contains non-whitespace text"));
                }
            },
            Event::Comment(_)
                if frames
                    .iter()
                    .any(|frame| matches!(frame, Frame::InnerExtensions | Frame::Opaque)) => {},
            Event::CData(_) if frames.contains(&Frame::Opaque) => {},
            Event::Eof => break,
            Event::Decl(_) if frames.is_empty() && !root_closed => {},
            Event::CData(_)
            | Event::Comment(_)
            | Event::DocType(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {
                return Err(invalid("Designer payload contains forbidden XML markup"));
            },
            _ => {},
        }
    }
    if !root_closed || !frames.is_empty() {
        return Err(invalid("Designer payload is unterminated"));
    }
    if root == Root::Tags && parsed.tags.is_none() {
        return Err(invalid("Designer tag list was not parsed"));
    }
    Ok(parsed)
}

fn finish_frame(
    xml: &[u8],
    frame: Frame,
    span: Range<usize>,
    parsed: &mut Parsed,
    inner_start: &mut Option<usize>,
) -> Result<()> {
    if frame == Frame::InnerExtensions {
        let bytes = xml
            .get(span)
            .ok_or_else(|| invalid("Designer inner extension range is invalid"))?;
        let mut retained = Vec::new();
        retained
            .try_reserve_exact(bytes.len())
            .map_err(|source| Error::Allocation {
                resource: "Designer inner extension",
                source,
            })?;
        retained.extend_from_slice(bytes);
        parsed.inner_extensions = Some(retained);
        *inner_start = None;
    }
    Ok(())
}

fn parse_editable(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<Option<bool>> {
    let mut editable = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = attribute.key.as_ref();
        if name == b"edtDesignElem" {
            if editable.is_some() {
                return Err(invalid("designPr has duplicate edtDesignElem attributes"));
            }
            if attribute.value.len() > limits.attribute_bytes() {
                return Err(limit("Designer attribute bytes", limits.attribute_bytes()));
            }
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?;
            if value.len() > limits.attribute_bytes() {
                return Err(limit("Designer attribute bytes", limits.attribute_bytes()));
            }
            editable = Some(parse_bool(&value)?);
        } else if name != b"xmlns" && !name.starts_with(b"xmlns:") {
            return Err(invalid("designPr has an unsupported attribute"));
        }
    }
    Ok(editable)
}

fn parse_tag(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<Tag> {
    let mut name = None;
    let mut value = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = attribute.key.as_ref();
        if key != b"name" && key != b"val" {
            if key == b"xmlns" || key.starts_with(b"xmlns:") {
                continue;
            }
            return Err(invalid(
                "designTag has an unsupported or qualified attribute",
            ));
        }
        if attribute.value.len() > limits.attribute_bytes() {
            return Err(limit("Designer attribute bytes", limits.attribute_bytes()));
        }
        let decoded = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        if decoded.len() > limits.attribute_bytes() {
            return Err(limit("Designer attribute bytes", limits.attribute_bytes()));
        }
        let slot = if key == b"name" {
            &mut name
        } else {
            &mut value
        };
        if slot.replace(decoded.into_owned()).is_some() {
            return Err(invalid("designTag has a duplicate required attribute"));
        }
    }
    Tag::new_with_limits(
        name.ok_or_else(|| invalid("designTag is missing name"))?,
        value.ok_or_else(|| invalid("designTag is missing val"))?,
        limits,
    )
}

fn validate_no_attributes(element: &BytesStart<'_>, limits: Limits) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = attribute.key.as_ref();
        if key != b"xmlns" && !key.starts_with(b"xmlns:") {
            return Err(invalid(
                "Designer list element has an unsupported attribute",
            ));
        }
        if attribute.value.len() > limits.attribute_bytes() {
            return Err(limit("Designer attribute bytes", limits.attribute_bytes()));
        }
    }
    Ok(())
}

fn is_p202(namespace: ResolveResult<'_>, proven_prefixes: Option<&[u8]>) -> bool {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => value == P202,
        ResolveResult::Unknown(prefix) => prefix_is_proven(proven_prefixes, prefix.as_slice()),
        ResolveResult::Unbound => prefix_is_proven(proven_prefixes, b""),
    }
}

fn prefix_is_proven(proof: Option<&[u8]>, prefix: &[u8]) -> bool {
    proof.is_some_and(|proof| proof.split(|byte| *byte == 0).any(|item| item == prefix))
}

fn parse_bool(value: &str) -> Result<bool> {
    match value.trim() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid("edtDesignElem is not an xsd:boolean")),
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_err| invalid("Designer XML offset does not fit usize"))
}

struct Output {
    bytes: Vec<u8>,
    limits: Limits,
}

impl Output {
    fn new(limits: Limits) -> Self {
        Self {
            bytes: Vec::new(),
            limits,
        }
    }

    fn push(&mut self, value: &[u8]) -> Result<()> {
        let length = self
            .bytes
            .len()
            .checked_add(value.len())
            .ok_or_else(|| limit("Designer output bytes", self.limits.xml_bytes()))?;
        if length > self.limits.xml_bytes() {
            return Err(limit("Designer output bytes", self.limits.xml_bytes()));
        }
        self.bytes
            .try_reserve(value.len())
            .map_err(|source| Error::Allocation {
                resource: "Designer output",
                source,
            })?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn push_escaped(&mut self, value: &str) -> Result<()> {
        for character in value.chars() {
            match character {
                '&' => self.push(b"&amp;")?,
                '<' => self.push(b"&lt;")?,
                '"' => self.push(b"&quot;")?,
                '\t' => self.push(b"&#x9;")?,
                '\n' => self.push(b"&#xA;")?,
                '\r' => self.push(b"&#xD;")?,
                _ => {
                    let mut encoded = [0u8; 4];
                    self.push(character.encode_utf8(&mut encoded).as_bytes())?;
                },
            }
        }
        Ok(())
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn properties_preserve_absence_empty_tags_duplicates_and_inner_extensions() {
        let xml = br#"<x:designPr xmlns:x="http://schemas.microsoft.com/office/powerpoint/2020/02/main" edtDesignElem="1"><x:designTagLst><x:designTag name="" val="same"/><x:designTag name="" val="same"/></x:designTagLst><x:extLst/></x:designPr>"#;
        let source = read_properties(xml, Limits::default()).unwrap();
        assert_eq!(source.value.editable(), Some(true));
        assert_eq!(source.value.tags().unwrap().len(), 2);
        assert_eq!(
            source.inner_extensions.as_deref(),
            Some(b"<x:extLst/>".as_slice())
        );

        // Canonical authoring does not relocate opaque inner XML because its
        // prefixes can depend on declarations inherited from the source.
        let written = write_properties(&source.value, None, Limits::default()).unwrap();
        let reparsed = read_properties(&written, Limits::default()).unwrap();
        assert_eq!(reparsed.value, source.value);
    }

    #[test]
    fn tags_escape_attributes_and_round_trip_control_whitespace() {
        let mut tags = Tags::new();
        tags.push(Tag::new("a&\"<", "\t\n\r").unwrap()).unwrap();
        let xml = write_tags(&tags, Limits::default()).unwrap();
        assert!(xml.windows(5).any(|value| value == b"&amp;"));
        assert_eq!(read_tags(&xml, Limits::default()).unwrap(), tags);
    }

    #[test]
    fn rejects_wrong_namespace_order_missing_attributes_and_forbidden_markup() {
        let wrong = br#"<p202:designTagLst xmlns:p202="urn:evil"/>"#;
        assert!(read_tags(wrong, Limits::default()).is_err());
        let order = br#"<p202:designPr xmlns:p202="http://schemas.microsoft.com/office/powerpoint/2020/02/main"><p202:extLst/><p202:designTagLst/></p202:designPr>"#;
        assert!(read_properties(order, Limits::default()).is_err());
        let missing = br#"<p202:designTagLst xmlns:p202="http://schemas.microsoft.com/office/powerpoint/2020/02/main"><p202:designTag name="x"/></p202:designTagLst>"#;
        assert!(read_tags(missing, Limits::default()).is_err());
        let dtd = br#"<!DOCTYPE x><p202:designTagLst xmlns:p202="http://schemas.microsoft.com/office/powerpoint/2020/02/main"/>"#;
        assert!(read_tags(dtd, Limits::default()).is_err());
    }

    #[test]
    fn exact_and_over_limits_are_enforced() {
        let xml = br#"<p202:designTagLst xmlns:p202="http://schemas.microsoft.com/office/powerpoint/2020/02/main"/>"#;
        let exact = Limits::default().with_xml_bytes(xml.len());
        assert!(read_tags(xml, exact).is_ok());
        assert!(read_tags(xml, exact.with_xml_bytes(xml.len() - 1)).is_err());
        assert!(read_tags(xml, exact.with_xml_nodes(0)).is_err());
    }

    #[test]
    fn detached_owner_proof_accepts_aliases_and_default_but_not_rebinding() {
        let aliases = br#"<a:designTagLst><b:designTag name="n" val="v"/></a:designTagLst>"#;
        let parsed = read_tags_with_prefix(aliases, Limits::default(), Some(b"a\0b")).unwrap();
        assert_eq!(parsed.as_slice()[0].value(), "v");

        let default = br#"<designTagLst><designTag name="n" val="v"/></designTagLst>"#;
        assert!(read_tags_with_prefix(default, Limits::default(), Some(b"")).is_ok());

        let rebound = br#"<a:designTagLst xmlns:b="urn:evil"><b:designTag name="n" val="v"/></a:designTagLst>"#;
        assert!(read_tags_with_prefix(rebound, Limits::default(), Some(b"a\0b")).is_err());
        assert!(read_tags_with_prefix(aliases, Limits::default(), Some(b"a")).is_err());
    }

    #[test]
    fn inner_extension_retains_opaque_character_content_but_rejects_active_markup() {
        let xml = br#"<p202:designPr xmlns:p202="http://schemas.microsoft.com/office/powerpoint/2020/02/main"><p202:extLst><p202:ext uri="u">text<!--comment--><![CDATA[<opaque>]]></p202:ext></p202:extLst></p202:designPr>"#;
        let parsed = read_properties(xml, Limits::default()).unwrap();
        assert_eq!(
            parsed.inner_extensions.as_deref(),
            Some(br#"<p202:extLst><p202:ext uri="u">text<!--comment--><![CDATA[<opaque>]]></p202:ext></p202:extLst>"#.as_slice())
        );

        let entity = br#"<p202:designPr xmlns:p202="http://schemas.microsoft.com/office/powerpoint/2020/02/main"><p202:extLst><p202:ext uri="u">&unsafe;</p202:ext></p202:extLst></p202:designPr>"#;
        assert!(read_properties(entity, Limits::default()).is_err());
        let pi = br#"<p202:designPr xmlns:p202="http://schemas.microsoft.com/office/powerpoint/2020/02/main"><p202:extLst><p202:ext uri="u"><?run x?></p202:ext></p202:extLst></p202:designPr>"#;
        assert!(read_properties(pi, Limits::default()).is_err());
    }
}
