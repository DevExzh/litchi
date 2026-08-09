use super::super::model::NamespaceDeclaration;
use super::super::{MAX_BYTES, MAX_DEPTH, MAX_NODES, MAX_STRING_BYTES};
use crate::{Error, Result};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

#[derive(Debug, Clone)]
pub(super) struct Fragment {
    pub namespace: String,
    pub local: String,
    pub attributes: Vec<(String, String)>,
    pub xml: Vec<u8>,
}

#[derive(Debug)]
pub(super) struct Scan {
    pub root: Fragment,
    pub children: Vec<Fragment>,
    pub namespaces: Vec<NamespaceDeclaration>,
}

#[derive(Debug)]
struct Open {
    namespace: String,
    local: String,
    start: usize,
    attributes: Vec<(String, String)>,
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn scan(xml: &[u8], label: &str) -> Result<Scan> {
    if xml.len() > MAX_BYTES {
        return Err(limit(label));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Open> = Vec::new();
    let mut root: Option<Fragment> = None;
    let mut children = Vec::new();
    let mut root_namespaces = Vec::new();
    let mut nodes = 0usize;
    let mut root_closed = false;

    loop {
        let start = reader.buffer_position() as usize;
        let decoder = reader.decoder();
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = resolve_namespace(resolved)?;
        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| limit(label))?;
                if nodes > MAX_NODES || stack.len() + 1 > MAX_DEPTH {
                    return Err(limit(label));
                }
                if root_closed && stack.is_empty() {
                    return Err(invalid(format!("{label} contains multiple roots")));
                }
                let open = Open {
                    namespace,
                    local: local_name(&element)?,
                    start,
                    attributes: attributes(&element, decoder)?,
                };
                if stack.is_empty() {
                    root_namespaces = namespace_declarations(&element, decoder)?;
                }
                stack.push(open);
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| limit(label))?;
                if nodes > MAX_NODES || stack.len() + 1 > MAX_DEPTH {
                    return Err(limit(label));
                }
                if root_closed && stack.is_empty() {
                    return Err(invalid(format!("{label} contains multiple roots")));
                }
                let fragment = Fragment {
                    namespace,
                    local: local_name(&element)?,
                    attributes: attributes(&element, decoder)?,
                    xml: xml[start..reader.buffer_position() as usize].to_vec(),
                };
                if stack.is_empty() {
                    root_namespaces = namespace_declarations(&element, decoder)?;
                    root = Some(fragment);
                    root_closed = true;
                } else if stack.len() == 1 {
                    children.push(fragment);
                }
            },
            Event::End(element) => {
                let open = stack
                    .pop()
                    .ok_or_else(|| invalid(format!("unexpected {label} closing element")))?;
                let local = end_local_name(&element)?;
                if open.namespace != namespace || open.local != local {
                    return Err(invalid(format!("mismatched {label} element")));
                }
                let end = reader.buffer_position() as usize;
                if stack.is_empty() {
                    root = Some(Fragment {
                        namespace: open.namespace,
                        local: open.local,
                        attributes: open.attributes,
                        xml: xml[open.start..end].to_vec(),
                    });
                    root_closed = true;
                } else if stack.len() == 1 {
                    children.push(Fragment {
                        namespace: open.namespace,
                        local: open.local,
                        attributes: open.attributes,
                        xml: xml[open.start..end].to_vec(),
                    });
                }
            },
            Event::Text(text) => {
                let value = text.decode().map_err(xml_error)?;
                let value = quick_xml::escape::unescape(&value).map_err(xml_error)?;
                if !value.trim().is_empty() && stack.is_empty() {
                    return Err(invalid(format!("unexpected {label} text")));
                }
            },
            Event::CData(text) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() && stack.is_empty() {
                    return Err(invalid(format!("unexpected {label} CDATA")));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(invalid(format!("DTD or processing instruction in {label}")));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() || !root_closed {
        return Err(invalid(format!("unterminated {label}")));
    }
    Ok(Scan {
        root: root.ok_or_else(|| invalid(format!("missing {label} root")))?,
        children,
        namespaces: root_namespaces,
    })
}

/// Parse a fragment with namespace declarations inherited from its owning
/// element. Captured child slices intentionally exclude those declarations so
/// the original opaque bytes remain untouched; this helper supplies the
/// context only while decoding known semantic children.
pub(super) fn scan_with_context(
    xml: &[u8],
    label: &str,
    context: &[NamespaceDeclaration],
) -> Result<Scan> {
    if context.is_empty() {
        return scan(xml, label);
    }
    let mut injected = xml.to_vec();
    let start =
        first_element_start(&injected).ok_or_else(|| invalid(format!("missing {label} root")))?;
    let end = injected[start..]
        .iter()
        .position(|byte| *byte == b'>')
        .map(|offset| start + offset)
        .ok_or_else(|| invalid(format!("unterminated {label} root")))?;
    let opening = injected[start..=end].to_vec();
    let mut insert = Vec::new();
    for declaration in context {
        let needle = if declaration.prefix.is_empty() {
            b"xmlns=".to_vec()
        } else {
            format!("xmlns:{}=", declaration.prefix).into_bytes()
        };
        if opening.windows(needle.len()).any(|window| window == needle) {
            continue;
        }
        insert.extend_from_slice(b" xmlns");
        if !declaration.prefix.is_empty() {
            insert.push(b':');
            insert.extend_from_slice(declaration.prefix.as_bytes());
        }
        insert.extend_from_slice(b"=\"");
        escape(&mut insert, &declaration.uri);
        insert.push(b'\"');
    }
    if !insert.is_empty() {
        let insertion = if injected[end.saturating_sub(1)] == b'/' {
            end - 1
        } else {
            end
        };
        injected.splice(insertion..insertion, insert);
    }
    scan(&injected, label)
}

fn first_element_start(xml: &[u8]) -> Option<usize> {
    let mut cursor = 0usize;
    while let Some(relative) = xml[cursor..].iter().position(|byte| *byte == b'<') {
        let start = cursor + relative;
        let next = xml.get(start + 1).copied();
        if !matches!(next, Some(b'?' | b'!')) {
            return Some(start);
        }
        cursor = start + 1;
    }
    None
}

pub(super) fn namespace_declarations(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Vec<NamespaceDeclaration>> {
    let mut output = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
        let Some(prefix) = key
            .strip_prefix("xmlns:")
            .or_else(|| key.eq("xmlns").then_some(""))
        else {
            continue;
        };
        let uri = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if uri.len() > MAX_STRING_BYTES {
            return Err(limit("modern comment namespace URI"));
        }
        output.push(NamespaceDeclaration {
            prefix: prefix.to_owned(),
            uri,
        });
    }
    Ok(output)
}

pub(super) fn attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Vec<(String, String)>> {
    let mut output = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
        if key == "xmlns" || key.starts_with("xmlns:") {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if key.len() > MAX_STRING_BYTES || value.len() > MAX_STRING_BYTES {
            return Err(limit("modern comment XML attribute"));
        }
        output.push((key.to_owned(), value));
    }
    Ok(output)
}

pub(super) fn local_name(element: &BytesStart<'_>) -> Result<String> {
    String::from_utf8(element.local_name().as_ref().to_vec()).map_err(xml_error)
}

pub(super) fn end_local_name(element: &BytesEnd<'_>) -> Result<String> {
    String::from_utf8(element.local_name().as_ref().to_vec()).map_err(xml_error)
}

pub(super) fn attribute<'a>(
    attributes: &'a [(String, String)],
    name: &str,
    required: bool,
) -> Result<Option<&'a str>> {
    let mut found = None;
    for (key, value) in attributes {
        if key == name {
            if found.is_some() {
                return Err(invalid(format!("duplicate XML attribute '{name}'")));
            }
            found = Some(value.as_str());
        } else if key.contains(':') {
            return Err(invalid(format!(
                "unexpected qualified XML attribute '{key}'"
            )));
        }
    }
    if required && found.is_none() {
        return Err(invalid(format!("missing required XML attribute '{name}'")));
    }
    Ok(found)
}

pub(super) fn no_attributes(attributes: &[(String, String)], label: &str) -> Result<()> {
    if let Some((name, _)) = attributes.first() {
        return Err(invalid(format!("unexpected {label} attribute '{name}'")));
    }
    Ok(())
}

pub(super) fn only_attributes(
    attributes: &[(String, String)],
    allowed: &[&str],
    label: &str,
) -> Result<()> {
    for (name, _) in attributes {
        if !allowed.contains(&name.as_str()) {
            return Err(invalid(format!("unexpected {label} attribute '{name}'")));
        }
    }
    Ok(())
}

pub(super) fn open(out: &mut Vec<u8>, prefix: &str, local: &str) {
    out.push(b'<');
    if !prefix.is_empty() {
        out.extend_from_slice(prefix.as_bytes());
        out.push(b':');
    }
    out.extend_from_slice(local.as_bytes());
}

pub(super) fn close(out: &mut Vec<u8>, prefix: &str, local: &str) {
    out.extend_from_slice(b"</");
    if !prefix.is_empty() {
        out.extend_from_slice(prefix.as_bytes());
        out.push(b':');
    }
    out.extend_from_slice(local.as_bytes());
    out.push(b'>');
}

pub(super) fn attr(out: &mut Vec<u8>, name: &str, value: &str) {
    out.push(b' ');
    out.extend_from_slice(name.as_bytes());
    out.extend_from_slice(b"=\"");
    escape(out, value);
    out.push(b'\"');
}

pub(super) fn namespaces(out: &mut Vec<u8>, declarations: &[NamespaceDeclaration]) {
    for declaration in declarations {
        out.extend_from_slice(b" xmlns");
        if !declaration.prefix.is_empty() {
            out.push(b':');
            out.extend_from_slice(declaration.prefix.as_bytes());
        }
        out.extend_from_slice(b"=\"");
        escape(out, &declaration.uri);
        out.push(b'\"');
    }
}

pub(super) fn escape(out: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '>' => out.extend_from_slice(b"&gt;"),
            '"' => out.extend_from_slice(b"&quot;"),
            '\'' => out.extend_from_slice(b"&apos;"),
            '\t' => out.extend_from_slice(b"&#x9;"),
            '\n' => out.extend_from_slice(b"&#xA;"),
            '\r' => out.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                out.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

pub(super) fn resolve_namespace(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            String::from_utf8(value.to_vec()).map_err(xml_error)
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(_) => {
            Err(invalid("modern comment XML uses an unresolved namespace"))
        },
    }
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    invalid(format!("invalid modern comment XML: {error}"))
}

pub(super) fn limit(label: &str) -> Error {
    invalid(format!("{label} exceeds implementation limit"))
}
