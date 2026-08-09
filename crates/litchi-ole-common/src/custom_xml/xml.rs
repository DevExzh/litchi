use std::borrow::Cow;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Limits, Result, RootName, invalid, limit, xml_error};

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
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid(
                    "DTD and general entity references are forbidden in Item XML",
                ));
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
        .chunks_exact(2)
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
) -> Result<String> {
    let mut value = None;
    for attribute_result in element.attributes().with_checks(true) {
        let attribute = attribute_result.map_err(|error| xml_error(error.to_string()))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() == local_name
            && matches!(resolved, ResolveResult::Bound(bound) if bound.as_ref() == namespace.as_bytes())
        {
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| xml_error(error.to_string()))?
                .into_owned();
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
) -> Result<()> {
    // The typed projection owns only the schema-defined attributes. Other
    // well-formed attributes are opaque extension data and remain in the
    // retained Properties stream; `allowed` documents the known fields and
    // keeps duplicate-required-attribute checks at their call sites.
    let _ = allowed;
    for attribute_result in element.attributes().with_checks(true) {
        let attribute = attribute_result.map_err(|error| xml_error(error.to_string()))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, _) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(resolved, ResolveResult::Unknown(_)) {
            return Err(invalid("XML attribute uses an unknown namespace prefix"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| xml_error(error.to_string()))?;
        validate_characters(&value)?;
    }
    Ok(())
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
