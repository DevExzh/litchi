use std::borrow::Cow;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Limits, Result, RootName, invalid, limit, xml_error};
use crate::xml_attributes::BytesStartExt as _;

#[derive(Clone, Copy)]
enum Utf16Encoding {
    LittleEndian,
    BigEndian,
}

pub(crate) fn validate_payload(xml: &[u8], limits: &Limits) -> Result<RootName> {
    if xml.is_empty() || xml.len() > limits.max_item_bytes {
        return Err(limit("Item XML is empty or exceeds its byte limit"));
    }
    let normalized = normalize_encoding(xml)?;
    let mut reader = NsReader::from_reader(normalized.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut root_name = None;
    let mut elements = 0usize;
    loop {
        buffer.clear();
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| xml_error(error.to_string()))?
        {
            Event::Start(element) => {
                elements += 1;
                if elements > limits.max_xml_elements {
                    return Err(limit("Item XML element count exceeds its limit"));
                }
                if depth == 0 {
                    roots += 1;
                    root_name = Some(expanded_root(&reader, &element)?);
                }
                depth += 1;
                if depth > limits.max_xml_depth {
                    return Err(limit("Item XML depth exceeds its limit"));
                }
            },
            Event::Empty(element) => {
                elements += 1;
                if elements > limits.max_xml_elements {
                    return Err(limit("Item XML element count exceeds its limit"));
                }
                if depth == 0 {
                    roots += 1;
                    root_name = Some(expanded_root(&reader, &element)?);
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("Item XML has an unexpected closing tag"))?;
            },
            Event::Text(text) if depth == 0 && !is_whitespace(text.as_ref()) => {
                return Err(invalid("Item XML has text outside its root"));
            },
            Event::CData(text) if depth == 0 && !is_whitespace(text.as_ref()) => {
                return Err(invalid("Item XML has CDATA outside its root"));
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "DTD and general entity references are forbidden in Item XML",
                ));
            },
            Event::GeneralRef(reference) => {
                if depth == 0 {
                    return Err(invalid("Item XML has a reference outside its root"));
                }
                validate_reference(&reference)?;
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_) => {},
        }
    }
    if roots != 1 || depth != 0 {
        return Err(invalid("Item XML must have exactly one complete root"));
    }
    root_name.ok_or_else(|| invalid("Item XML has no root"))
}

/// Validate a numeric or predefined reference without resolving application
/// entities or changing the retained lexical bytes. Return whether it is XML
/// whitespace, for callers whose known grammar permits no other text.
pub(crate) fn validate_reference(reference: &BytesRef<'_>) -> Result<bool> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| xml_error(error.to_string()))?
    {
        let mut encoded = [0; 4];
        let value = character.encode_utf8(&mut encoded);
        validate_characters(value)?;
        return Ok(is_whitespace(value.as_bytes()));
    }
    let name = reference
        .decode()
        .map_err(|error| xml_error(error.to_string()))?;
    if quick_xml::escape::resolve_predefined_entity(&name).is_none() {
        return Err(invalid(
            "general entity references are forbidden in Custom XML",
        ));
    }
    Ok(false)
}

fn expanded_root(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<RootName> {
    let (resolved, local_name_bytes) = reader.resolver().resolve_element(element.name());
    let namespace = match resolved {
        ResolveResult::Bound(Namespace(value)) => Some(
            std::str::from_utf8(value)
                .map_err(|_utf8_error| xml_error("root namespace is not UTF-8"))?
                .to_string(),
        ),
        ResolveResult::Unbound => None,
        ResolveResult::Unknown(prefix) => {
            return Err(xml_error(format!(
                "root uses unknown namespace prefix {prefix:?}"
            )));
        },
    };
    let local_name = std::str::from_utf8(local_name_bytes.as_ref())
        .map_err(|_utf8_error| xml_error("root local name is not UTF-8"))?
        .to_string();
    Ok(RootName {
        namespace,
        local_name,
    })
}

pub(crate) fn normalize_encoding(xml: &[u8]) -> Result<Cow<'_, [u8]>> {
    let (encoding_hint, bytes) = if let Some(bytes) = xml.strip_prefix(&[0xFF, 0xFE]) {
        (Some(Utf16Encoding::LittleEndian), bytes)
    } else if let Some(bytes) = xml.strip_prefix(&[0xFE, 0xFF]) {
        (Some(Utf16Encoding::BigEndian), bytes)
    } else if xml.starts_with(&[b'<', 0, b'?', 0])
        || xml.starts_with(&[b'<', 0, b'!', 0])
        || xml.starts_with(&[b'<', 0])
    {
        (Some(Utf16Encoding::LittleEndian), xml)
    } else if xml.starts_with(&[0, b'<', 0, b'?'])
        || xml.starts_with(&[0, b'<', 0, b'!'])
        || xml.starts_with(&[0, b'<'])
    {
        (Some(Utf16Encoding::BigEndian), xml)
    } else {
        (None, xml)
    };
    let Some(utf16_encoding) = encoding_hint else {
        let text = std::str::from_utf8(xml)
            .map_err(|_utf8_error| xml_error("XML is not valid UTF-8 or UTF-16"))?;
        validate_characters(text)?;
        return Ok(Cow::Borrowed(xml));
    };
    if !bytes.len().is_multiple_of(2) {
        return Err(xml_error("UTF-16 XML has an odd byte length"));
    }
    let units = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| match utf16_encoding {
            Utf16Encoding::LittleEndian => u16::from_le_bytes([pair[0], pair[1]]),
            Utf16Encoding::BigEndian => u16::from_be_bytes([pair[0], pair[1]]),
        })
        .collect::<Vec<_>>();
    let text = String::from_utf16(&units)
        .map_err(|_utf16_error| xml_error("UTF-16 XML is not well-formed Unicode"))?;
    validate_characters(&text)?;
    Ok(Cow::Owned(text.into_bytes()))
}

/// Re-encode a normalized Properties document using the source document's
/// UTF-8/UTF-16 representation.  Property rewrites operate on normalized
/// UTF-8 so that source offsets are stable, then restore the original wire
/// encoding before publishing the retained stream.
pub(crate) fn restore_encoding(
    source: &[u8],
    normalized: &[u8],
    maximum: usize,
) -> Result<Vec<u8>> {
    let (little_endian, bom) = if source.starts_with(&[0xFF, 0xFE]) {
        (true, true)
    } else if source.starts_with(&[0xFE, 0xFF]) {
        (false, true)
    } else if source.starts_with(&[b'<', 0, b'?', 0]) || source.starts_with(&[b'<', 0]) {
        (true, false)
    } else if source.starts_with(&[0, b'<', 0, b'?']) || source.starts_with(&[0, b'<']) {
        (false, false)
    } else {
        if normalized.len() > maximum {
            return Err(limit("serialized Properties XML exceeds its byte limit"));
        }
        return Ok(normalized.to_vec());
    };

    let text = std::str::from_utf8(normalized)
        .map_err(|error| xml_error(format!("normalized XML is not UTF-8: {error}")))?;
    let body_len = text
        .encode_utf16()
        .count()
        .checked_mul(2)
        .ok_or_else(|| limit("UTF-16 Properties XML size overflows usize"))?;
    let output_len = body_len
        .checked_add(if bom { 2 } else { 0 })
        .ok_or_else(|| limit("UTF-16 Properties XML size overflows usize"))?;
    if output_len > maximum {
        return Err(limit("serialized Properties XML exceeds its byte limit"));
    }
    let mut output = Vec::with_capacity(output_len);
    if bom {
        output.extend_from_slice(if little_endian {
            &[0xFF, 0xFE]
        } else {
            &[0xFE, 0xFF]
        });
    }
    for unit in text.encode_utf16() {
        let bytes = if little_endian {
            unit.to_le_bytes()
        } else {
            unit.to_be_bytes()
        };
        output.extend_from_slice(&bytes);
    }
    Ok(output)
}

pub(crate) fn validate_characters(value: &str) -> Result<()> {
    if value.chars().any(|character| {
        !matches!(
            character,
            '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(xml_error("XML contains a character forbidden by XML 1.0"));
    }
    Ok(())
}

pub(crate) fn resolved_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<(Option<Vec<u8>>, Vec<u8>)> {
    let (resolved, local) = reader.resolver().resolve_element(element.name());
    let namespace = match resolved {
        ResolveResult::Bound(namespace) => Some(namespace.as_ref().to_vec()),
        ResolveResult::Unbound => None,
        ResolveResult::Unknown(prefix) => {
            return Err(xml_error(format!(
                "unknown element namespace prefix {prefix:?}"
            )));
        },
    };
    Ok((namespace, local.as_ref().to_vec()))
}

pub(crate) fn required_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    local_name: &[u8],
    limits: &Limits,
) -> Result<String> {
    let mut value = None;
    for attribute_result in element.checked_attributes() {
        let attribute = attribute_result.map_err(|error| xml_error(error.to_string()))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() == local_name
            && matches!(resolved, ResolveResult::Bound(bound) if bound.as_ref() == namespace.as_bytes())
        {
            preflight_attribute_value(&attribute, limits)?;
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| xml_error(error.to_string()))?;
            if decoded.len() > limits.max_string_bytes {
                return Err(limit("XML attribute value exceeds its byte limit"));
            }
            let decoded = decoded.into_owned();
            if value.replace(decoded).is_some() {
                return Err(invalid("XML element has a duplicate attribute"));
            }
        }
    }
    value.ok_or_else(|| {
        invalid(format!(
            "XML element lacks required attribute {}",
            String::from_utf8_lossy(local_name)
        ))
    })
}

pub(crate) fn reject_other_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    allowed: &[(&str, &[u8])],
    limits: &Limits,
) -> Result<()> {
    // The typed projection owns only the schema-defined attributes. Other
    // well-formed attributes are opaque extension data and remain in the
    // retained Properties stream; `allowed` documents the known fields and
    // keeps duplicate-required-attribute checks at their call sites.
    let _ = allowed;
    for attribute_result in element.checked_attributes() {
        let attribute = attribute_result.map_err(|error| xml_error(error.to_string()))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, _) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(resolved, ResolveResult::Unknown(_)) {
            return Err(invalid("XML attribute uses an unknown namespace prefix"));
        }
        preflight_attribute_value(&attribute, limits)?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| xml_error(error.to_string()))?;
        validate_characters(&value)?;
    }
    Ok(())
}

fn preflight_attribute_value(
    attribute: &quick_xml::events::attributes::Attribute<'_>,
    limits: &Limits,
) -> Result<()> {
    let length = normalized_attribute_len(attribute.value.as_ref())?;
    if length > limits.max_string_bytes {
        return Err(limit("XML attribute value exceeds its byte limit"));
    }
    Ok(())
}

/// Compute the exact XML 1.0 normalized, decoded attribute length without
/// constructing the decoded string.  `Properties` parsing normalizes source
/// encoding before quick-xml sees it, so this scan receives UTF-8 bytes just
/// like the decoder used by `decoded_and_normalized_value`.
fn normalized_attribute_len(value: &[u8]) -> Result<usize> {
    let value = std::str::from_utf8(value)
        .map_err(|error| xml_error(format!("XML attribute is not UTF-8: {error}")))?;
    let bytes = value.as_bytes();
    let mut cursor = 0usize;
    let mut length = 0usize;
    while cursor < bytes.len() {
        match bytes[cursor] {
            b'\t' | b'\n' => {
                length = length
                    .checked_add(1)
                    .ok_or_else(|| limit("XML attribute length overflows usize"))?;
                cursor += 1;
            },
            b'\r' => {
                length = length
                    .checked_add(1)
                    .ok_or_else(|| limit("XML attribute length overflows usize"))?;
                cursor += 1;
                if bytes.get(cursor) == Some(&b'\n') {
                    cursor += 1;
                }
            },
            b'&' => {
                let entity_start = cursor + 1;
                let relative_end = bytes
                    .get(entity_start..)
                    .and_then(|remaining| remaining.iter().position(|byte| *byte == b';'))
                    .ok_or_else(|| xml_error("XML attribute has an unterminated entity"))?;
                let entity_end = entity_start + relative_end;
                let entity = &value[entity_start..entity_end];
                let decoded_length = if let Some(number) = entity.strip_prefix('#') {
                    let code = if let Some(hex) = number.strip_prefix('x') {
                        u32::from_str_radix(hex, 16)
                    } else {
                        number.parse::<u32>()
                    }
                    .map_err(|error| {
                        xml_error(format!("invalid XML character reference: {error}"))
                    })?;
                    let character = char::from_u32(code).ok_or_else(|| {
                        xml_error("XML character reference is not a scalar value")
                    })?;
                    character.len_utf8()
                } else {
                    quick_xml::escape::resolve_predefined_entity(entity)
                        .ok_or_else(|| xml_error("XML attribute uses an unknown entity"))?
                        .len()
                };
                length = length
                    .checked_add(decoded_length)
                    .ok_or_else(|| limit("XML attribute length overflows usize"))?;
                cursor = entity_end + 1;
            },
            _ => {
                let character = value[cursor..]
                    .chars()
                    .next()
                    .ok_or_else(|| xml_error("XML attribute contains invalid UTF-8"))?;
                length = length
                    .checked_add(character.len_utf8())
                    .ok_or_else(|| limit("XML attribute length overflows usize"))?;
                cursor += character.len_utf8();
            },
        }
    }
    Ok(length)
}

pub(crate) fn escape_attribute(output: &mut Vec<u8>, value: &str) {
    for byte in value.bytes() {
        match byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'"' => output.extend_from_slice(b"&quot;"),
            b'\t' => output.extend_from_slice(b"&#x9;"),
            b'\n' => output.extend_from_slice(b"&#xA;"),
            b'\r' => output.extend_from_slice(b"&#xD;"),
            _ => output.push(byte),
        }
    }
}

pub(crate) fn is_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}
