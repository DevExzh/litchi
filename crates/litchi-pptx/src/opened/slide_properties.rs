//! Bounded, lexical edits for the two slide-root properties owned by this
//! crate: `p:cSld/@name` and `p:sld/@show`.
//!
//! The editor keeps every byte outside those attributes intact.  In
//! particular, unknown attributes and extension payloads are never
//! reconstructed from a semantic model.  The same rooted scan supplies the
//! source-patch proof skeleton used by the deferred editor.

use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use litchi_opc::xml_attributes::BytesStartExt;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use sha2::{Digest, Sha256};

use super::model::invalid;
use crate::{Error, Result};

const PROPERTY_RESOURCE: &str = "opened-presentation slide properties XML";

#[derive(Debug, Clone, PartialEq, Eq)]
struct RawAttribute {
    name: Range<usize>,
    value: Range<usize>,
    full: Range<usize>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TagKind {
    Slide,
    CommonSlide,
}

#[derive(Debug, Clone)]
struct Tag {
    range: Range<usize>,
    attributes: Vec<RawAttribute>,
    kind: TagKind,
}

#[derive(Debug)]
struct Document {
    slide: Tag,
    common_slide: Tag,
    slide_name: Option<String>,
    show: bool,
}

/// Stage one or both slide-root property updates.
///
/// `name` updates the authored `p:cSld/@name` attribute. `hidden` maps to the
/// inverse of `p:sld/@show`; an omitted `show` attribute is therefore already
/// semantically visible. An exact semantic no-op returns `Ok(None)` and keeps
/// the original bytes untouched.
pub(crate) fn stage_slide_properties(
    xml: &[u8],
    name: Option<&str>,
    hidden: Option<bool>,
    max_output_bytes: usize,
) -> Result<Option<Vec<u8>>> {
    stage_rooted_properties(xml, b"sld", name, hidden, max_output_bytes)
}

/// Stage a `p:cSld/@name` update of a slide master (`root` `sldMaster`) or
/// slide layout (`sldLayout`) part, with the same lexical guarantees as a
/// slide's (record 0321). An equal name returns `Ok(None)`.
pub(crate) fn stage_common_slide_name(
    xml: &[u8],
    root: &[u8],
    name: &str,
    max_output_bytes: usize,
) -> Result<Option<Vec<u8>>> {
    stage_rooted_properties(xml, root, Some(name), None, max_output_bytes)
}

fn stage_rooted_properties(
    xml: &[u8],
    root: &[u8],
    name: Option<&str>,
    hidden: Option<bool>,
    max_output_bytes: usize,
) -> Result<Option<Vec<u8>>> {
    let document = parse_rooted(xml, root)?;
    if let Some(value) = name
        && !value.chars().all(is_xml_char)
    {
        return Err(invalid(
            "opened-presentation slide name contains an invalid XML character",
        ));
    }

    let name_changed = name.is_some_and(|value| document.slide_name.as_deref() != Some(value));
    let show_changed = hidden.is_some_and(|value| document.show == value);
    if !name_changed && !show_changed {
        return Ok(None);
    }

    let escaped_name = if name_changed {
        Some(escape_attribute(
            name.unwrap_or_default(),
            max_output_bytes,
        )?)
    } else {
        None
    };
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(usize::from(name_changed) + usize::from(show_changed))
        .map_err(|source| Error::Allocation {
            resource: PROPERTY_RESOURCE,
            source,
        })?;
    if name_changed {
        add_edit(
            xml,
            &document.common_slide,
            b"name",
            escaped_name.as_deref().unwrap_or_default(),
            &mut edits,
        )?;
    }
    if show_changed {
        add_edit(
            xml,
            &document.slide,
            b"show",
            if hidden == Some(false) { b"1" } else { b"0" },
            &mut edits,
        )?;
    }

    let mut output_len = xml.len();
    for edit in &edits {
        output_len = output_len
            .checked_sub(edit.span.end - edit.span.start)
            .and_then(|length| length.checked_add(edit.value.len()))
            .ok_or(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
                observed: usize::MAX,
                maximum: max_output_bytes,
            })?;
    }
    if output_len > max_output_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
            observed: output_len,
            maximum: max_output_bytes,
        });
    }
    edits.sort_unstable_by_key(|edit| edit.span.start);
    if edits
        .windows(2)
        .any(|pair| pair[0].span.end > pair[1].span.start)
    {
        return Err(invalid("opened-presentation slide property edits overlap"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: PROPERTY_RESOURCE,
            source,
        })?;
    let mut cursor = 0;
    for edit in edits {
        output.extend_from_slice(&xml[cursor..edit.span.start]);
        output.extend_from_slice(&edit.value);
        cursor = edit.span.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    if output.len() != output_len {
        return Err(invalid(
            "opened-presentation slide property output length mismatch",
        ));
    }

    // Read the candidate back through the same rooted parser before
    // publishing it. This also proves that a newly inserted attribute uses
    // the source's namespace/root topology and XML lexical rules.
    let staged = parse_rooted(&output, root)?;
    if name_changed && staged.slide_name.as_deref() != name {
        return Err(invalid(
            "opened-presentation slide name did not round-trip semantically",
        ));
    }
    if show_changed && staged.show == hidden.unwrap_or(false) {
        return Err(invalid(
            "opened-presentation slide visibility did not round-trip semantically",
        ));
    }
    Ok(Some(output))
}

/// Fingerprint all slide XML bytes except the two managed attributes.
///
/// Presence, lexical spelling, entity spelling, and value width of `name` or
/// `show` are intentionally erased. All unknown attributes and every other
/// byte remain part of the proof surface.
pub(crate) fn slide_properties_skeleton(xml: &[u8]) -> Result<[u8; 32]> {
    let document = parse_document(xml)?;
    let mut tags = [document.slide, document.common_slide];
    tags.sort_unstable_by_key(|tag| tag.range.start);
    let mut hash = Sha256::new();
    hash.update(b"litchi-pptx-slide-properties-skeleton-v1\0");
    let mut cursor = 0;
    for tag in &tags {
        hash.update(&xml[cursor..tag.range.start]);
        hash_tag_without_managed_attributes(&mut hash, xml, tag)?;
        cursor = tag.range.end;
    }
    hash.update(&xml[cursor..]);
    Ok(hash.finalize().into())
}

/// Remove the managed slide attributes while retaining every other source
/// byte. The rooted parser runs first, so malformed values, duplicate managed
/// attributes, and invalid root topology cannot be hidden by normalization.
/// Each removed span includes only its leading XML whitespace and the complete
/// managed attribute; unknown attributes and their lexical spacing remain.
pub(crate) fn normalize_slide_properties(xml: &[u8]) -> Result<Vec<u8>> {
    let document = parse_document(xml)?;
    let mut ranges = Vec::new();
    ranges
        .try_reserve_exact(2)
        .map_err(|source| Error::Allocation {
            resource: PROPERTY_RESOURCE,
            source,
        })?;
    if let Some(range) = managed_attribute_range(xml, &document.slide, b"show") {
        ranges.push(range);
    }
    if let Some(range) = managed_attribute_range(xml, &document.common_slide, b"name") {
        ranges.push(range);
    }
    ranges.sort_unstable_by_key(|range| range.start);
    if ranges.windows(2).any(|pair| pair[0].end > pair[1].start) {
        return Err(invalid(
            "opened-presentation slide property normalization spans overlap",
        ));
    }
    let removed = ranges.iter().try_fold(0usize, |total, range| {
        total
            .checked_add(range.end - range.start)
            .ok_or_else(|| invalid("opened-presentation slide property span overflows"))
    })?;
    let output_len = xml
        .len()
        .checked_sub(removed)
        .ok_or_else(|| invalid("opened-presentation slide property output underflows"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: PROPERTY_RESOURCE,
            source,
        })?;
    let mut cursor = 0;
    for range in ranges {
        output.extend_from_slice(&xml[cursor..range.start]);
        cursor = range.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    if output.len() != output_len {
        return Err(invalid(
            "opened-presentation slide property normalization length mismatch",
        ));
    }
    Ok(output)
}

#[derive(Debug)]
struct Edit {
    span: Range<usize>,
    value: Vec<u8>,
}

fn add_edit(
    xml: &[u8],
    tag: &Tag,
    attribute: &[u8],
    value: &[u8],
    edits: &mut Vec<Edit>,
) -> Result<()> {
    if let Some(raw) = tag.attributes.iter().find(|raw| {
        &xml[tag.range.start + raw.name.start..tag.range.start + raw.name.end] == attribute
    }) {
        let mut replacement = Vec::new();
        replacement
            .try_reserve_exact(value.len())
            .map_err(|source| Error::Allocation {
                resource: PROPERTY_RESOURCE,
                source,
            })?;
        replacement.extend_from_slice(value);
        edits.push(Edit {
            span: (tag.range.start + raw.value.start)..(tag.range.start + raw.value.end),
            value: replacement,
        });
        return Ok(());
    }
    let insertion = tag
        .range
        .end
        .checked_sub(1)
        .ok_or_else(|| invalid("opened-presentation slide property tag is empty"))?;
    let insertion = if xml.get(insertion.wrapping_sub(1)) == Some(&b'/') {
        insertion - 1
    } else {
        insertion
    };
    let mut replacement = Vec::new();
    let required = attribute
        .len()
        .checked_add(value.len())
        .and_then(|length| length.checked_add(4))
        .ok_or_else(|| invalid("opened-presentation slide property length overflows"))?;
    replacement
        .try_reserve_exact(required)
        .map_err(|source| Error::Allocation {
            resource: PROPERTY_RESOURCE,
            source,
        })?;
    replacement.extend_from_slice(b" ");
    replacement.extend_from_slice(attribute);
    replacement.extend_from_slice(b"=\"");
    replacement.extend_from_slice(value);
    replacement.extend_from_slice(b"\"");
    edits.push(Edit {
        span: insertion..insertion,
        value: replacement,
    });
    Ok(())
}

fn managed_attribute_range(xml: &[u8], tag: &Tag, attribute: &[u8]) -> Option<Range<usize>> {
    tag.attributes.iter().find_map(|raw| {
        (&xml[tag.range.start + raw.name.start..tag.range.start + raw.name.end] == attribute)
            .then(|| (tag.range.start + raw.full.start)..(tag.range.start + raw.full.end))
    })
}

fn escape_attribute(value: &str, maximum: usize) -> Result<Vec<u8>> {
    fn escaped_length(character: char) -> usize {
        match character {
            '&' => 5,
            '<' | '>' => 4,
            '\'' | '"' => 6,
            '\t' => 4,
            '\n' | '\r' => 5,
            _ => character.len_utf8(),
        }
    }
    let length = value.chars().try_fold(0usize, |total, character| {
        total
            .checked_add(escaped_length(character))
            .ok_or(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
                observed: usize::MAX,
                maximum,
            })
    })?;
    if length > maximum {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
            observed: length,
            maximum,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: PROPERTY_RESOURCE,
            source,
        })?;
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '\'' => output.extend_from_slice(b"&apos;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#9;"),
            '\n' => output.extend_from_slice(b"&#10;"),
            '\r' => output.extend_from_slice(b"&#13;"),
            character => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
    Ok(output)
}

fn parse_document(xml: &[u8]) -> Result<Document> {
    parse_rooted(xml, b"sld")
}

fn parse_rooted(xml: &[u8], root: &[u8]) -> Result<Document> {
    let origin = ReaderOrigin::of(xml);
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_count = 0usize;
    let mut namespace = None::<Vec<u8>>;
    let mut slide = None;
    let mut common_slide = None;
    let mut slide_name = None;
    let mut show = true;
    loop {
        let (resolved, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        // Keep only static known namespaces so the reader borrow ends before
        // inspecting its source position and decoder.
        let resolved = match resolved {
            ResolveResult::Bound(Namespace(value))
                if value == crate::namespace::PRESENTATIONML_NAMESPACE =>
            {
                ResolveResult::Bound(Namespace(crate::namespace::PRESENTATIONML_NAMESPACE))
            },
            ResolveResult::Bound(Namespace(value))
                if value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE =>
            {
                ResolveResult::Bound(Namespace(crate::namespace::STRICT_PRESENTATIONML_NAMESPACE))
            },
            _ => ResolveResult::Unbound,
        };
        let event_end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("slide property source offset overflows"))?;
        let starts = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let start = event_start(xml, event_end)?;
                let current_depth = depth;
                if current_depth == 0 {
                    root_count = root_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("opened-presentation slide root count overflows"))?;
                    if root_count != 1 || !is_pml(&resolved, element.local_name().as_ref(), root) {
                        return Err(invalid(
                            "opened-presentation slide properties require one p:sld root",
                        ));
                    }
                    let root_namespace = pml_namespace(&resolved)?;
                    namespace = Some(root_namespace.to_vec());
                    let tag = make_tag(xml, start, event_end, TagKind::Slide)?;
                    show = target_value(&element, b"show", reader.decoder())?
                        .map_or(Ok(true), |value| parse_bool(&value, "slide show"))?;
                    slide = Some(tag);
                } else if current_depth == 1
                    && is_pml(&resolved, element.local_name().as_ref(), b"cSld")
                {
                    let root_namespace = namespace.as_deref().ok_or_else(|| {
                        invalid("opened-presentation slide root namespace is missing")
                    })?;
                    if pml_namespace(&resolved)? != root_namespace {
                        return Err(invalid(
                            "opened-presentation slide cSld namespace differs from its root",
                        ));
                    }
                    if common_slide.is_some() {
                        return Err(invalid(
                            "opened-presentation slide contains duplicate p:cSld roots",
                        ));
                    }
                    let tag = make_tag(xml, start, event_end, TagKind::CommonSlide)?;
                    slide_name = target_value(&element, b"name", reader.decoder())?;
                    if slide_name
                        .as_deref()
                        .is_some_and(|value| !value.chars().all(is_xml_char))
                    {
                        return Err(invalid(
                            "opened-presentation slide name contains an invalid XML character",
                        ));
                    }
                    common_slide = Some(tag);
                }
                if starts {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("opened-presentation slide XML depth overflows"))?;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opened-presentation slide has an unexpected end"))?;
            },
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Text(text)
                if depth == 0 && text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::DocType(_) => {
                return Err(invalid(
                    "opened-presentation slide document types are not publishable",
                ));
            },
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) => {
                if depth == 0 {
                    return Err(invalid(
                        "opened-presentation slide has non-whitespace content outside its root",
                    ));
                }
            },
            Event::Eof => break,
        }
    }
    if depth != 0 || root_count != 1 {
        return Err(invalid(
            "opened-presentation slide root topology is incomplete",
        ));
    }
    let slide = slide.ok_or_else(|| invalid("opened-presentation slide root is missing"))?;
    let common_slide = common_slide
        .ok_or_else(|| invalid("opened-presentation slide lacks a direct p:cSld root"))?;
    Ok(Document {
        slide,
        common_slide,
        slide_name,
        show,
    })
}

fn make_tag(xml: &[u8], start: usize, end: usize, kind: TagKind) -> Result<Tag> {
    if start >= end || end > xml.len() || xml.get(start) != Some(&b'<') {
        return Err(invalid(
            "opened-presentation slide property tag span is invalid",
        ));
    }
    Ok(Tag {
        range: start..end,
        attributes: scan_attributes(&xml[start..end])?,
        kind,
    })
}

fn event_start(xml: &[u8], event_end: usize) -> Result<usize> {
    if event_end > xml.len() {
        return Err(invalid(
            "opened-presentation slide event exceeds its source",
        ));
    }
    xml[..event_end]
        .iter()
        .rposition(|byte| *byte == b'<')
        .ok_or_else(|| invalid("opened-presentation slide event has no opening delimiter"))
}

fn scan_attributes(tag: &[u8]) -> Result<Vec<RawAttribute>> {
    let mut attributes = Vec::new();
    let mut cursor = 1usize;
    while cursor < tag.len() && !is_name_delimiter(tag[cursor]) {
        cursor += 1;
    }
    while cursor < tag.len() {
        let full_start = cursor;
        while cursor < tag.len() && is_xml_space(tag[cursor]) {
            cursor += 1;
        }
        if cursor >= tag.len() || tag[cursor] == b'>' {
            break;
        }
        if tag[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !matches!(tag[cursor], b'=' | b'>' | b'/')
            && !is_xml_space(tag[cursor])
        {
            cursor += 1;
        }
        if name_start == cursor {
            return Err(invalid("opened-presentation slide attribute name is empty"));
        }
        let name_end = cursor;
        while cursor < tag.len() && is_xml_space(tag[cursor]) {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return Err(invalid(
                "opened-presentation slide attribute lacks an equals sign",
            ));
        }
        cursor += 1;
        while cursor < tag.len() && is_xml_space(tag[cursor]) {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .ok_or_else(|| invalid("opened-presentation slide attribute lacks a value"))?;
        if quote != b'"' && quote != b'\'' {
            return Err(invalid("opened-presentation slide attribute is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        if cursor >= tag.len() {
            return Err(invalid(
                "opened-presentation slide attribute is unterminated",
            ));
        }
        let value_end = cursor;
        cursor += 1;
        attributes
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: PROPERTY_RESOURCE,
                source,
            })?;
        attributes.push(RawAttribute {
            name: name_start..name_end,
            value: value_start..value_end,
            full: full_start..cursor,
        });
    }
    Ok(attributes)
}

fn target_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.value.contains(&b'<') {
            return Err(invalid("slide attributes contain a raw opening delimiter"));
        }
        let decoded = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() != name {
            continue;
        }
        if value.is_some() {
            return Err(invalid(format!(
                "opened-presentation slide property repeats `{}`",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(decoded.into_owned());
    }
    Ok(value)
}

fn parse_bool(value: &str, field: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!(
            "opened-presentation {field} must be 0, 1, true, or false"
        ))),
    }
}

fn pml_namespace<'a>(namespace: &'a ResolveResult<'a>) -> Result<&'a [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value))
            if *value == crate::namespace::PRESENTATIONML_NAMESPACE
                || *value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE =>
        {
            Ok(value)
        },
        _ => Err(invalid(
            "opened-presentation slide root has an invalid namespace",
        )),
    }
}

fn is_pml(namespace: &ResolveResult<'_>, local: &[u8], expected: &[u8]) -> bool {
    local == expected
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if *value == crate::namespace::PRESENTATIONML_NAMESPACE
                    || *value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE
        )
}

fn hash_tag_without_managed_attributes(hash: &mut Sha256, xml: &[u8], tag: &Tag) -> Result<()> {
    hash.update(&xml[tag.range.start..tag.range.start + 1]);
    let source = &xml[tag.range.clone()];
    let managed = match tag.kind {
        TagKind::Slide => b"show".as_slice(),
        TagKind::CommonSlide => b"name".as_slice(),
    };
    let mut cursor = 1usize;
    for attribute in &tag.attributes {
        let is_managed = &source[attribute.name.clone()] == managed;
        if !is_managed {
            hash.update(&source[cursor..attribute.full.start]);
            hash.update(&source[attribute.full.clone()]);
            cursor = attribute.full.end;
        } else {
            hash.update(&source[cursor..attribute.full.start]);
            cursor = attribute.full.end;
        }
    }
    hash.update(&source[cursor..]);
    Ok(())
}

fn is_xml_space(value: u8) -> bool {
    matches!(value, b' ' | b'\t' | b'\n' | b'\r')
}

fn is_name_delimiter(value: u8) -> bool {
    is_xml_space(value) || matches!(value, b'>' | b'/')
}

fn is_xml_char(value: char) -> bool {
    matches!(
        value,
        '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
    )
}

#[cfg(test)]
mod tests {
    use super::{normalize_slide_properties, slide_properties_skeleton, stage_slide_properties};

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

    fn source() -> Vec<u8> {
        format!(
            r#"<?xml version="1.0"?> <p:sld xmlns:p="{PML}" xmlns:x="urn:private" x:keep="yes" show="true"><p:cSld x:unknown="a&amp;b" name="Old &amp; Name"><p:spTree/></p:cSld><p:extLst><x:payload/></p:extLst></p:sld>"#
        )
        .into_bytes()
    }

    #[test]
    fn semantic_noops_preserve_lexical_bytes() {
        let xml = source();
        assert!(
            stage_slide_properties(&xml, Some("Old & Name"), Some(false), xml.len())
                .unwrap()
                .is_none()
        );
        assert!(
            stage_slide_properties(&xml, None, Some(false), xml.len())
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn changed_properties_preserve_unknown_bytes_and_entities() {
        let xml = source();
        let staged = stage_slide_properties(&xml, Some("New & Name"), Some(true), xml.len() + 64)
            .unwrap()
            .unwrap();
        assert!(String::from_utf8_lossy(&staged).contains("name=\"New &amp; Name\""));
        assert!(String::from_utf8_lossy(&staged).contains("show=\"0\""));
        assert!(String::from_utf8_lossy(&staged).contains("x:unknown=\"a&amp;b\""));
        assert_eq!(
            slide_properties_skeleton(&xml).unwrap(),
            slide_properties_skeleton(&staged).unwrap()
        );
    }

    #[test]
    fn name_whitespace_uses_character_references() {
        let xml = source();
        let staged = stage_slide_properties(
            &xml,
            Some("line\ncolumn\ttab\rreturn"),
            None,
            xml.len() + 64,
        )
        .unwrap()
        .unwrap();
        let raw = String::from_utf8_lossy(&staged);
        assert!(raw.contains("line&#10;column&#9;tab&#13;return"));
    }

    #[test]
    fn omitted_show_is_visible_and_adds_only_when_hiding() {
        let xml =
            format!(r#"<p:sld xmlns:p="{PML}"><p:cSld><p:spTree/></p:cSld></p:sld>"#).into_bytes();
        assert!(
            stage_slide_properties(&xml, None, Some(false), xml.len())
                .unwrap()
                .is_none()
        );
        let staged = stage_slide_properties(&xml, None, Some(true), xml.len() + 16)
            .unwrap()
            .unwrap();
        assert!(String::from_utf8_lossy(&staged).contains("show=\"0\""));
    }

    #[test]
    fn normalization_removes_only_managed_attributes() {
        let xml = source();
        let normalized = normalize_slide_properties(&xml).unwrap();
        let normalized = String::from_utf8_lossy(&normalized);
        assert!(!normalized.contains("show=\"true\""));
        assert!(!normalized.contains("name=\"Old &amp; Name\""));
        assert!(normalized.contains("x:keep=\"yes\""));
        assert!(normalized.contains("x:unknown=\"a&amp;b\""));
        assert!(normalized.contains("<p:sld xmlns:p=\""));
    }

    #[test]
    fn malformed_roots_and_invalid_names_are_refused() {
        let xml = b"<p:sld xmlns:p=\"urn:wrong\"><p:cSld/></p:sld>";
        assert!(stage_slide_properties(xml, Some("x"), None, 4096).is_err());
        let xml = source();
        assert!(stage_slide_properties(&xml, Some("bad\u{1}"), None, 4096).is_err());
        let xml = format!(r#"<p:sld xmlns:p="{PML}"><p:cSld/></p:wrong>"#);
        assert!(stage_slide_properties(xml.as_bytes(), Some("x"), None, 4096).is_err());
    }
}
