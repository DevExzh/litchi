//! Strict decoding, lossless splicing and canonical encoding of the
//! extended-properties part.

use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::{Field, MAX_TEXT_BYTES, MAX_XML_BYTES, NAMESPACE, Props, STRICT_NAMESPACE, Value};
use crate::xml::decode_xml_reference;
use crate::{Error, Result};

const MAX_EVENTS: usize = 262_144;
const MAX_DEPTH: usize = 64;
const XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

/// Where one scalar element sits in the part.
#[derive(Debug, Clone)]
struct Element {
    field: Field,
    /// The whole element, start tag through end tag.
    span: Range<usize>,
    /// The content between the tags; `None` for an empty element (`<X/>`).
    content: Option<Range<usize>>,
    /// The element's qualified name as written.
    name: Vec<u8>,
}

/// The decoded part: its typed value and where everything sits.
#[derive(Debug)]
struct Layout {
    props: Props,
    elements: Vec<Element>,
    /// The root's qualified name as written.
    root_name: Vec<u8>,
    /// For an empty root (`<Properties/>`), the span of its start tag;
    /// otherwise the offset of its end tag.
    root_end: RootEnd,
}

#[derive(Debug, Clone)]
enum RootEnd {
    Empty(Range<usize>),
    EndTag(usize),
}

/// Decode an extended-properties part strictly (see the module
/// documentation).
///
/// # Errors
///
/// Returns [`Error::Invalid`] for a structural violation or a value outside
/// its lexical space, [`Error::Xml`] for malformed XML, and [`Error::Limit`]
/// when a bound is exceeded.
pub(crate) fn decode(xml: &[u8]) -> Result<Props> {
    Ok(scan(xml)?.props)
}

/// Rewrite `xml` so it decodes to `target`, touching only the scalar
/// elements whose value changes. Returns `None` when `xml` already decodes
/// to `target`.
///
/// A changed element keeps its start tag (prefix and attributes) and gets
/// new content; a removed element is cut out with its tags; a new element
/// is appended at the end of the root, named with the root's prefix. The
/// result is decoded again and must equal `target`.
///
/// # Errors
///
/// The errors of [`decode`] for `xml`, and [`Error::Invalid`] when a target
/// text value holds a character XML 1.0 cannot carry.
pub(crate) fn splice(xml: &[u8], target: &Props) -> Result<Option<Vec<u8>>> {
    let layout = scan(xml)?;
    if layout.props == *target {
        return Ok(None);
    }
    let prefix = qualified_prefix(&layout.root_name);
    let mut edits: Vec<(Range<usize>, Vec<u8>)> = Vec::new();
    let mut appended = Vec::new();
    for field in Field::ALL {
        let before = field.get(&layout.props);
        let after = field.get(target);
        if before == after {
            continue;
        }
        let element = layout
            .elements
            .iter()
            .find(|element| element.field == field);
        match (element, after) {
            (Some(element), None) => edits.push((element.span.clone(), Vec::new())),
            (Some(element), Some(value)) => {
                let text = lexical(field, value)?;
                match &element.content {
                    Some(content) => edits.push((content.clone(), text.into_bytes())),
                    None => {
                        // `<X .../>`: keep the start tag, replace `/>`.
                        let end = element.span.end;
                        let mut replacement =
                            Vec::with_capacity(text.len() + element.name.len() + 4);
                        replacement.push(b'>');
                        replacement.extend_from_slice(text.as_bytes());
                        replacement.extend_from_slice(b"</");
                        replacement.extend_from_slice(&element.name);
                        replacement.push(b'>');
                        edits.push((end - 2..end, replacement));
                    },
                }
            },
            (None, Some(value)) => {
                push_element(&mut appended, prefix, field, value)?;
            },
            (None, None) => {},
        }
    }
    if !appended.is_empty() {
        match &layout.root_end {
            RootEnd::EndTag(offset) => edits.push((*offset..*offset, appended)),
            RootEnd::Empty(span) => {
                let mut replacement =
                    Vec::with_capacity(appended.len() + layout.root_name.len() + 4);
                replacement.push(b'>');
                replacement.extend_from_slice(&appended);
                replacement.extend_from_slice(b"</");
                replacement.extend_from_slice(&layout.root_name);
                replacement.push(b'>');
                edits.push((span.end - 2..span.end, replacement));
            },
        }
    }
    edits.sort_by_key(|(range, _)| (range.start, range.end));
    let mut output = Vec::with_capacity(xml.len().saturating_add(256));
    let mut cursor = 0;
    for (range, replacement) in edits {
        if range.start < cursor || range.end > xml.len() {
            return Err(Error::Invalid(
                "extended-properties edits overlap".to_owned(),
            ));
        }
        output.extend_from_slice(&xml[cursor..range.start]);
        output.extend_from_slice(&replacement);
        cursor = range.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    if output.len() > MAX_XML_BYTES {
        return Err(Error::Limit {
            resource: "extended-properties XML bytes",
            max: MAX_XML_BYTES,
            actual: output.len(),
        });
    }
    if decode(&output)? != *target {
        return Err(Error::Invalid(
            "extended-properties edit did not read back as requested".to_owned(),
        ));
    }
    Ok(Some(output))
}

/// Encode a new Transitional part holding `props`, elements in schema order.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when a text value holds a character XML 1.0
/// cannot carry, and [`Error::Limit`] when a bound is exceeded.
pub(crate) fn encode_new(props: &Props) -> Result<Vec<u8>> {
    let mut xml = Vec::with_capacity(512);
    xml.extend_from_slice(XML_DECLARATION.as_bytes());
    xml.extend_from_slice(b"<Properties xmlns=\"");
    xml.extend_from_slice(NAMESPACE.as_bytes());
    xml.extend_from_slice(b"\" xmlns:vt=\"");
    xml.extend_from_slice(super::VT_NAMESPACE.as_bytes());
    xml.extend_from_slice(b"\"");
    let mut children = Vec::new();
    for field in Field::ALL {
        if let Some(value) = field.get(props) {
            push_element(&mut children, &[], field, value)?;
        }
    }
    if children.is_empty() {
        xml.extend_from_slice(b"/>");
    } else {
        xml.push(b'>');
        xml.extend_from_slice(&children);
        xml.extend_from_slice(b"</Properties>");
    }
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::Limit {
            resource: "extended-properties XML bytes",
            max: MAX_XML_BYTES,
            actual: xml.len(),
        });
    }
    Ok(xml)
}

fn qualified_prefix(name: &[u8]) -> &[u8] {
    name.iter()
        .position(|byte| *byte == b':')
        .map_or(&[][..], |colon| &name[..colon])
}

fn push_element(output: &mut Vec<u8>, prefix: &[u8], field: Field, value: Value<'_>) -> Result<()> {
    let text = lexical(field, value)?;
    let name = field.local_name().as_bytes();
    let qualified = |output: &mut Vec<u8>| {
        if !prefix.is_empty() {
            output.extend_from_slice(prefix);
            output.push(b':');
        }
        output.extend_from_slice(name);
    };
    output.push(b'<');
    qualified(output);
    if text.is_empty() {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.push(b'>');
    output.extend_from_slice(text.as_bytes());
    output.extend_from_slice(b"</");
    qualified(output);
    output.push(b'>');
    Ok(())
}

/// The escaped lexical form of one value.
fn lexical(field: Field, value: Value<'_>) -> Result<String> {
    match value {
        Value::Text(text) => {
            if text.len() > MAX_TEXT_BYTES {
                return Err(Error::Limit {
                    resource: "extended-property text bytes",
                    max: MAX_TEXT_BYTES,
                    actual: text.len(),
                });
            }
            if let Some(character) = text.chars().find(|character| !is_xml_char(*character)) {
                return Err(Error::Invalid(format!(
                    "extended property {} contains XML 1.0-forbidden character U+{:04X}",
                    field.local_name(),
                    u32::from(character)
                )));
            }
            let mut escaped = String::with_capacity(text.len());
            for character in text.chars() {
                match character {
                    '&' => escaped.push_str("&amp;"),
                    '<' => escaped.push_str("&lt;"),
                    '>' => escaped.push_str("&gt;"),
                    // A literal carriage return would read back as a line
                    // feed (XML end-of-line handling).
                    '\r' => escaped.push_str("&#13;"),
                    other => escaped.push(other),
                }
            }
            Ok(escaped)
        },
        Value::Int(value) => Ok(value.to_string()),
        Value::Bool(value) => Ok(if value { "true" } else { "false" }.to_owned()),
    }
}

fn is_xml_char(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(u32::from(character), 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

/// One reader event, detached from the reader's buffer.
enum Scanned {
    /// A start tag (`empty` for `<x/>`): whether it is in the root's
    /// namespace (or, for the root, in an extended-properties namespace),
    /// its local name and its qualified name.
    Element {
        empty: bool,
        namespace: Option<Vec<u8>>,
        local: Vec<u8>,
        name: Vec<u8>,
    },
    End,
    Text(String),
    Other,
    DocType,
    Eof,
}

fn scan(xml: &[u8]) -> Result<Layout> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::Limit {
            resource: "extended-properties XML bytes",
            max: MAX_XML_BYTES,
            actual: xml.len(),
        });
    }
    if crate::mce::mentions_namespace(xml) {
        return Err(Error::Invalid(
            "extended properties that use markup compatibility are not supported".to_owned(),
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let origin = ReaderOrigin::of(xml);
    let mut props = Props::new();
    let mut elements: Vec<Element> = Vec::new();
    // The root's qualified name and namespace once read.
    let mut root: Option<(Vec<u8>, Vec<u8>)> = None;
    let mut root_end = None;
    let mut depth = 0usize;
    // The scalar element being read: field, start offset, start-tag end,
    // qualified name, decoded text.
    let mut active: Option<(Field, usize, usize, Vec<u8>, String)> = None;
    let mut events = 0usize;

    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(Error::Limit {
                resource: "extended-properties XML events",
                max: MAX_EVENTS,
                actual: events,
            });
        }
        let start = offset(origin, reader.buffer_position())?;
        let scanned = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(format!("extended properties XML error: {error}")))?;
            let namespace = match namespace {
                ResolveResult::Bound(Namespace(value)) => Some(value.to_vec()),
                ResolveResult::Unbound => None,
                ResolveResult::Unknown(prefix) => {
                    return Err(Error::Invalid(format!(
                        "extended properties use the undeclared prefix '{}'",
                        String::from_utf8_lossy(&prefix)
                    )));
                },
            };
            match event {
                Event::Start(element) => Scanned::Element {
                    empty: false,
                    namespace,
                    local: element.local_name().as_ref().to_vec(),
                    name: element.name().as_ref().to_vec(),
                },
                Event::Empty(element) => Scanned::Element {
                    empty: true,
                    namespace,
                    local: element.local_name().as_ref().to_vec(),
                    name: element.name().as_ref().to_vec(),
                },
                Event::End(_) => Scanned::End,
                Event::Text(text) => {
                    let decoded = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    Scanned::Text(
                        quick_xml::escape::unescape(&decoded)
                            .map_err(|error| Error::Xml(error.to_string()))?
                            .into_owned(),
                    )
                },
                Event::CData(text) => Scanned::Text(
                    text.xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .into_owned(),
                ),
                Event::GeneralRef(reference) => Scanned::Text(decode_xml_reference(&reference)?),
                Event::DocType(_) => Scanned::DocType,
                Event::Eof => Scanned::Eof,
                Event::Comment(_) | Event::Decl(_) | Event::PI(_) => Scanned::Other,
            }
        };
        let end = offset(origin, reader.buffer_position())?;
        match scanned {
            Scanned::Element {
                empty,
                namespace,
                local,
                name,
            } if root.is_none() => {
                let Some(namespace) = namespace.filter(|value| {
                    value.as_slice() == NAMESPACE.as_bytes()
                        || value.as_slice() == STRICT_NAMESPACE.as_bytes()
                }) else {
                    return Err(Error::Invalid(
                        "expected Properties in the extended-properties namespace".to_owned(),
                    ));
                };
                if local != b"Properties" {
                    return Err(Error::Invalid(
                        "expected Properties in the extended-properties namespace".to_owned(),
                    ));
                }
                if empty {
                    root_end = Some(RootEnd::Empty(start..end));
                } else {
                    depth = 1;
                }
                root = Some((name, namespace));
            },
            Scanned::Element {
                empty,
                namespace,
                local,
                name,
            } => {
                if root_end.is_some() {
                    return Err(Error::Invalid(
                        "Properties must be the only document element".to_owned(),
                    ));
                }
                if active.is_some() {
                    return Err(Error::Invalid(
                        "an extended property holds element content".to_owned(),
                    ));
                }
                let field = (depth == 1
                    && namespace.as_deref() == root.as_ref().map(|(_, root)| root.as_slice()))
                .then(|| Field::from_local_name(&local))
                .flatten();
                if let Some(field) = field {
                    if elements.iter().any(|element| element.field == field) {
                        return Err(Error::Invalid(format!(
                            "extended property {} is repeated",
                            field.local_name()
                        )));
                    }
                    if empty {
                        field.set_lexical(&mut props, String::new())?;
                        elements.push(Element {
                            field,
                            span: start..end,
                            content: None,
                            name,
                        });
                    } else {
                        active = Some((field, start, end, name, String::new()));
                        depth += 1;
                    }
                } else if !empty {
                    depth += 1;
                    if depth > MAX_DEPTH {
                        return Err(Error::Limit {
                            resource: "extended-properties XML depth",
                            max: MAX_DEPTH,
                            actual: depth,
                        });
                    }
                }
            },
            Scanned::End => {
                if let Some((field, element_start, content_start, name, text)) = active.take() {
                    field.set_lexical(&mut props, text)?;
                    elements.push(Element {
                        field,
                        span: element_start..end,
                        content: Some(content_start..start),
                        name,
                    });
                    depth -= 1;
                    continue;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("unexpected end tag in extended properties".to_owned())
                })?;
                if depth == 0 {
                    root_end = Some(RootEnd::EndTag(start));
                }
            },
            Scanned::Text(text) => {
                if let Some((.., value)) = active.as_mut() {
                    push_text(value, &text)?;
                } else if depth <= 1 && !text.trim_matches([' ', '\t', '\n', '\r']).is_empty() {
                    return Err(Error::Invalid(
                        "extended properties hold text outside a property".to_owned(),
                    ));
                }
            },
            Scanned::DocType => {
                return Err(Error::Invalid(
                    "DTD is not allowed in extended properties".to_owned(),
                ));
            },
            Scanned::Eof => break,
            Scanned::Other => {},
        }
    }
    let (Some((root_name, _)), Some(root_end)) = (root, root_end) else {
        return Err(Error::Invalid(
            "unterminated or missing Properties root element".to_owned(),
        ));
    };
    if depth != 0 || active.is_some() {
        return Err(Error::Invalid(
            "unterminated extended-properties element".to_owned(),
        ));
    }
    Ok(Layout {
        props,
        elements,
        root_name,
        root_end,
    })
}

fn push_text(value: &mut String, text: &str) -> Result<()> {
    if value.len().saturating_add(text.len()) > MAX_TEXT_BYTES {
        return Err(Error::Limit {
            resource: "extended-property text bytes",
            max: MAX_TEXT_BYTES,
            actual: value.len().saturating_add(text.len()),
        });
    }
    value.push_str(text);
    Ok(())
}

fn offset(origin: ReaderOrigin, position: u64) -> Result<usize> {
    origin
        .offset(position)
        .ok_or_else(|| Error::Invalid("extended-properties XML offset exceeds usize".to_owned()))
}
