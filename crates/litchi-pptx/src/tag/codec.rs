use super::model::*;
use super::package::process_pptx_ooxml;
use super::*;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

/// Parse one bounded Strict or Transitional tag-list part.
pub fn parse(xml: &[u8]) -> Result<List> {
    parse_profiled(xml).map(|(list, _)| list)
}

pub(crate) fn parse_profiled(xml: &[u8]) -> Result<(List, Conformance)> {
    if xml.len() > MAX_PART_BYTES {
        return Err(Error::Limit {
            resource: "tag-list bytes",
            limit: MAX_PART_BYTES,
        });
    }
    let xml = process_pptx_ooxml(xml, MAX_PART_BYTES)?;
    if xml.len() > MAX_PART_BYTES {
        return Err(Error::Limit {
            resource: "MCE-expanded tag-list bytes",
            limit: MAX_PART_BYTES,
        });
    }
    let mut reader = NsReader::from_reader(xml.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root = false;
    let mut closed = false;
    let mut open_tag: Option<(usize, Tag)> = None;
    let mut conformance = None;
    let mut tags = Vec::new();
    let mut namespaces = Vec::new();
    let mut attrs = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                let name = element.local_name();
                if !root && depth == 0 && pml(&namespace).is_some() && name.as_ref() == b"tagLst" {
                    root = true;
                    conformance = pml(&namespace);
                    let parsed = parse_attributes(&element, &[], reader.decoder())?;
                    namespaces = parsed.namespaces;
                    attrs = parsed.extensions;
                    depth = 1;
                } else if root
                    && !closed
                    && depth == 1
                    && open_tag.is_none()
                    && pml(&namespace) == conformance
                    && name.as_ref() == b"tag"
                {
                    if tags.len() == MAX_TAGS {
                        return Err(Error::Limit {
                            resource: "tag count",
                            limit: MAX_TAGS,
                        });
                    }
                    let tag = parse_tag(&element, reader.decoder())?;
                    depth = depth.saturating_add(1);
                    open_tag = Some((depth, tag));
                } else {
                    return Err(invalid(format!(
                        "unexpected tag-list element '{}'",
                        String::from_utf8_lossy(name.as_ref())
                    )));
                }
            },
            Event::Empty(element) => {
                let name = element.local_name();
                if !root && depth == 0 && pml(&namespace).is_some() && name.as_ref() == b"tagLst" {
                    conformance = pml(&namespace);
                    let parsed = parse_attributes(&element, &[], reader.decoder())?;
                    namespaces = parsed.namespaces;
                    attrs = parsed.extensions;
                    root = true;
                    closed = true;
                } else if root
                    && !closed
                    && depth == 1
                    && open_tag.is_none()
                    && pml(&namespace) == conformance
                    && name.as_ref() == b"tag"
                {
                    if tags.len() == MAX_TAGS {
                        return Err(Error::Limit {
                            resource: "tag count",
                            limit: MAX_TAGS,
                        });
                    }
                    tags.push(parse_tag(&element, reader.decoder())?);
                } else {
                    return Err(invalid(format!(
                        "unexpected tag-list element '{}'",
                        String::from_utf8_lossy(name.as_ref())
                    )));
                }
            },
            Event::End(element) => {
                let name = element.local_name();
                if open_tag.as_ref().is_some_and(|(level, _)| *level == depth)
                    && pml(&namespace) == conformance
                    && name.as_ref() == b"tag"
                {
                    let Some((_, tag)) = open_tag.take() else {
                        return Err(invalid("tag parser state is inconsistent"));
                    };
                    tags.push(tag);
                    depth = depth.saturating_sub(1);
                } else if root
                    && !closed
                    && depth == 1
                    && pml(&namespace) == conformance
                    && name.as_ref() == b"tagLst"
                {
                    closed = true;
                    depth = 0;
                } else {
                    return Err(invalid("unexpected tag-list end element"));
                }
            },
            Event::Text(text) => {
                let value = text.decode().map_err(xml_error)?;
                let value = quick_xml::escape::unescape(&value).map_err(xml_error)?;
                if !value.trim().is_empty() {
                    return Err(invalid("tag elements cannot contain text"));
                }
            },
            Event::CData(text) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("tag elements cannot contain CDATA"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DTDs and processing instructions are rejected in tag lists",
                ));
            },
            Event::GeneralRef(_) => {
                return Err(invalid("tag elements cannot contain entity references"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !root || !closed || depth != 0 || open_tag.is_some() {
        return Err(invalid("unterminated tag-list part"));
    }
    let wire_len = list_wire_len_parts(&tags, &namespaces, &attrs)?;
    let list = List {
        tags,
        namespaces,
        attrs,
        wire_len,
    };
    validate_structure(&list)?;
    let conformance =
        conformance.ok_or_else(|| invalid("tag-list namespace profile is missing"))?;
    Ok((list, conformance))
}

/// Encode one detached list without interpreting any retained value.
pub fn write(value: &List, conformance: Conformance) -> Result<Vec<u8>> {
    validate_structure(value)?;
    validate_unique_names(value)?;
    ensure_list_budget(value.wire_len)?;
    let mut out = Vec::new();
    out.try_reserve_exact(value.wire_len)
        .map_err(|source| allocation("encoded tag-list output", source))?;
    append(&mut out, XML_DECL)?;
    append(&mut out, ROOT_OPEN)?;
    escape(&mut out, conformance.namespace())?;
    push(&mut out, b'\"')?;
    for attr in &value.namespaces {
        write_preserved(&mut out, attr)?;
    }
    for attr in &value.attrs {
        write_preserved(&mut out, attr)?;
    }
    if value.tags.is_empty() {
        append(&mut out, ROOT_EMPTY_CLOSE)?;
        return Ok(out);
    }
    append(&mut out, ROOT_CHILDREN_OPEN)?;
    for tag in &value.tags {
        append(&mut out, TAG_OPEN)?;
        for attr in &tag.namespaces {
            write_preserved(&mut out, attr)?;
        }
        for attr in &tag.attrs {
            write_preserved(&mut out, attr)?;
        }
        write_attr(&mut out, "name", &tag.name)?;
        write_attr(&mut out, "val", &tag.value)?;
        append(&mut out, TAG_CLOSE)?;
    }
    append(&mut out, ROOT_CLOSE)?;
    Ok(out)
}

struct ParsedAttributes {
    values: Vec<(String, String)>,
    namespaces: Vec<raw::Attr>,
    extensions: Vec<raw::Attr>,
}

fn parse_attributes(
    element: &BytesStart<'_>,
    known: &[&str],
    decoder: Decoder,
) -> Result<ParsedAttributes> {
    let mut values = Vec::new();
    let mut namespaces = Vec::new();
    let mut extensions = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(xml_error)?
            .to_string();
        validate_qname(&name)?;
        if !seen.insert(name.clone()) {
            return Err(invalid("duplicate XML attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bounded_text(&value, "tag attribute")?;
        if name == "xmlns:p" {
            if !matches!(value.as_str(), PML_TEXT | STRICT_TEXT) {
                return Err(invalid(
                    "p prefix is bound to a non-PresentationML namespace",
                ));
            }
        } else if is_namespace(&name) {
            namespaces.push(raw::Attr {
                qualified_name: name,
                value,
            });
        } else if !name.contains(':') && known.contains(&name.as_str()) {
            values.push((name, value));
        } else if name.contains(':') {
            extensions.push(raw::Attr {
                qualified_name: name,
                value,
            });
        } else {
            return Err(invalid(format!("unexpected tag attribute '{name}'")));
        }
    }
    Ok(ParsedAttributes {
        values,
        namespaces,
        extensions,
    })
}

fn parse_tag(element: &BytesStart<'_>, decoder: Decoder) -> Result<Tag> {
    let parsed = parse_attributes(element, &["name", "val"], decoder)?;
    let mut name = None;
    let mut value = None;
    for (key, item) in parsed.values {
        match key.as_str() {
            "name" => name = Some(item),
            "val" => value = Some(item),
            _ => return Err(invalid("unexpected parsed tag attribute")),
        }
    }
    let name = name.ok_or_else(|| invalid("tag is missing 'name'"))?;
    let value = value.ok_or_else(|| invalid("tag is missing 'val'"))?;
    let wire_len = tag_wire_len_parts(&name, &value, &parsed.namespaces, &parsed.extensions)?;
    Ok(Tag {
        name,
        value,
        namespaces: parsed.namespaces,
        attrs: parsed.extensions,
        wire_len,
    })
}

fn write_preserved(out: &mut Vec<u8>, attr: &raw::Attr) -> Result<()> {
    validate_qname(attr.qualified_name())?;
    bounded_text(attr.value(), "tag attribute")?;
    write_attr(out, attr.qualified_name(), attr.value())
}

fn write_attr(out: &mut Vec<u8>, name: &str, value: &str) -> Result<()> {
    push(out, b' ')?;
    append(out, name.as_bytes())?;
    append(out, b"=\"")?;
    escape(out, value)?;
    push(out, b'\"')
}

fn escape(out: &mut Vec<u8>, value: &str) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => append(out, b"&amp;")?,
            '<' => append(out, b"&lt;")?,
            '"' => append(out, b"&quot;")?,
            '\t' => append(out, b"&#x9;")?,
            '\n' => append(out, b"&#xA;")?,
            '\r' => append(out, b"&#xD;")?,
            _ => {
                let mut encoded = [0; 4];
                append(out, character.encode_utf8(&mut encoded).as_bytes())?;
            },
        }
    }
    Ok(())
}

fn append(out: &mut Vec<u8>, bytes: &[u8]) -> Result<()> {
    let len = out
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| invalid("tag-list output length overflow"))?;
    if len > MAX_PART_BYTES {
        return Err(Error::Limit {
            resource: "encoded tag-list bytes",
            limit: MAX_PART_BYTES,
        });
    }
    out.extend_from_slice(bytes);
    Ok(())
}

fn push(out: &mut Vec<u8>, byte: u8) -> Result<()> {
    append(out, std::slice::from_ref(&byte))
}

pub(crate) fn is_namespace(value: &str) -> bool {
    value == "xmlns" || value.starts_with("xmlns:")
}

pub(crate) fn validate_qname(value: &str) -> Result<()> {
    let mut parts = value.split(':');
    let Some(first) = parts.next() else {
        return Err(invalid("tag attribute QName is empty"));
    };
    let second = parts.next();
    let valid = valid_ncname(first)
        && second.is_none_or(valid_ncname)
        && parts.next().is_none()
        && value.len() <= MAX_TEXT_BYTES;
    if valid {
        Ok(())
    } else {
        Err(invalid(format!("invalid tag attribute QName '{value}'")))
    }
}

fn valid_ncname(value: &str) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    (first == '_' || first.is_alphabetic())
        && chars.all(|character| {
            character == '_'
                || character == '-'
                || character == '.'
                || character == '\u{B7}'
                || character.is_alphanumeric()
                || matches!(character, '\u{0300}'..='\u{036F}' | '\u{203F}'..='\u{2040}')
        })
}

pub(crate) fn bounded_text(value: &str, resource: &'static str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(Error::Limit {
            resource,
            limit: MAX_TEXT_BYTES,
        });
    }
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(invalid(format!(
            "{resource} contains a character forbidden by XML 1.0"
        )))
    }
}

fn is_xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
}

pub(crate) fn pml(namespace: &ResolveResult<'_>) -> Option<Conformance> {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == PML => Some(Conformance::Transitional),
        ResolveResult::Bound(value) if value.as_ref() == STRICT => Some(Conformance::Strict),
        _ => None,
    }
}

pub(crate) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
