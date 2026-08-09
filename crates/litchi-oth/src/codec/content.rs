//! Bounded, namespace-aware ODF text-web content validation.

use litchi_core::{Error, Result};
use litchi_odf_common::compact_xml;
use quick_xml::{
    XmlVersion,
    events::Event,
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_DEPTH: usize = compact_xml::DEFAULT_MAX_DEPTH;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Element {
    DocumentContent,
    Body,
    Text,
    Paragraph,
    Other,
}

/// Validate a compact ODF text-web `content.xml` part.
///
/// The check intentionally validates only the package-family envelope, not
/// every ODF text grammar production. In particular, it requires exactly one
/// expanded-name `office:document-content/office:body/office:text` chain and
/// accepts namespace-prefix aliases.
///
/// # Errors
///
/// Returns an error for non-compact or malformed XML, DTDs or custom entity
/// references, an over-depth tree, or a wrong/misplaced text-web envelope.
pub(crate) fn validate(xml: &str) -> Result<()> {
    compact_xml::validate(xml.as_bytes()).map_err(Error::from)?;
    validate_structure(xml)
}

/// Project direct and inline character data from `text:p` elements.
///
/// This deliberately does not expand fields, evaluate links, fetch resources,
/// or synthesize list labels. It is a bounded read-only snapshot projection.
pub(crate) fn paragraphs(xml: &str) -> Result<Vec<crate::paragraph::Paragraph>> {
    validate(xml)?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::new();
    let mut values = Vec::new();
    let mut paragraph = None;

    loop {
        match reader.read_event().map_err(xml_error)? {
            Event::Start(start) => {
                let current = element(&reader, start.name());
                reserve(&mut stack, "OTH XML element stack")?;
                stack.push(current);
                if current == Element::Paragraph {
                    if paragraph.is_some() {
                        return invalid("OTH text paragraph cannot contain another text paragraph");
                    }
                    paragraph = Some(String::new());
                }
            },
            Event::Empty(start) => {
                if element(&reader, start.name()) == Element::Paragraph {
                    reserve(&mut values, "OTH paragraph projection")?;
                    values.push(crate::paragraph::Paragraph::new(String::new()));
                }
            },
            Event::Text(text) if paragraph.is_some() => {
                let text = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid OTH content.xml character data: {error}"))
                })?;
                append_text(paragraph.as_mut(), &text)?;
            },
            Event::CData(text) if paragraph.is_some() => {
                let text = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid OTH content.xml character data: {error}"))
                })?;
                append_text(paragraph.as_mut(), &text)?;
            },
            Event::End(end) => {
                let current = element(&reader, end.name());
                let Some(open) = stack.pop() else {
                    return invalid("OTH XML end tag has no matching start tag");
                };
                if current != open {
                    return invalid("OTH XML end tag does not match its start tag");
                }
                if current == Element::Paragraph {
                    let Some(text) = paragraph.take() else {
                        return invalid("OTH text paragraph has no active text projection");
                    };
                    reserve(&mut values, "OTH paragraph projection")?;
                    values.push(crate::paragraph::Paragraph::new(text));
                }
            },
            Event::Eof => return Ok(values),
            _ => {},
        }
    }
}

fn validate_structure(xml: &str) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::new();
    let mut body_seen = false;
    let mut text_seen = false;
    let mut root_closed = false;

    loop {
        match reader.read_event().map_err(xml_error)? {
            Event::Start(start) => {
                let current = element(&reader, start.name());
                handle_start(
                    &mut stack,
                    current,
                    &mut body_seen,
                    &mut text_seen,
                    root_closed,
                )?;
            },
            Event::Empty(start) => {
                let current = element(&reader, start.name());
                handle_empty(&stack, current, &mut body_seen, &mut text_seen, root_closed)?;
            },
            Event::End(end) => {
                let current = element(&reader, end.name());
                let Some(open) = stack.pop() else {
                    return invalid("OTH XML end tag has no matching start tag");
                };
                if current != open {
                    return invalid("OTH XML end tag does not match its start tag");
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                if !stack.contains(&Element::Paragraph) {
                    return invalid(
                        "OTH content.xml contains formatting whitespace outside text content",
                    );
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !root_closed || !stack.is_empty() {
        return invalid("OTH content.xml has no closed document-content root");
    }
    if !body_seen || !text_seen {
        return invalid("OTH content.xml must contain exactly one office:body/office:text chain");
    }
    Ok(())
}

fn handle_start(
    stack: &mut Vec<Element>,
    current: Element,
    body_seen: &mut bool,
    text_seen: &mut bool,
    root_closed: bool,
) -> Result<()> {
    if root_closed || stack.len() >= MAX_DEPTH {
        return invalid("OTH content.xml has an invalid element depth or trailing root");
    }
    if current == Element::Paragraph && !stack.contains(&Element::Text) {
        return invalid("OTH text:p must occur inside the office:text body");
    }
    validate_position(
        stack.last().copied(),
        stack.is_empty(),
        current,
        body_seen,
        text_seen,
    )?;
    reserve(stack, "OTH XML element stack")?;
    stack.push(current);
    Ok(())
}

fn handle_empty(
    stack: &[Element],
    current: Element,
    body_seen: &mut bool,
    text_seen: &mut bool,
    root_closed: bool,
) -> Result<()> {
    if root_closed || stack.len() >= MAX_DEPTH {
        return invalid("OTH content.xml has an invalid empty element depth or trailing root");
    }
    if current == Element::Paragraph && !stack.contains(&Element::Text) {
        return invalid("OTH text:p must occur inside the office:text body");
    }
    validate_position(
        stack.last().copied(),
        stack.is_empty(),
        current,
        body_seen,
        text_seen,
    )?;
    match current {
        Element::DocumentContent | Element::Body => {
            invalid("OTH content.xml cannot use an empty root or body element")
        },
        Element::Text | Element::Paragraph | Element::Other => Ok(()),
    }
}

fn validate_position(
    parent: Option<Element>,
    root: bool,
    current: Element,
    body_seen: &mut bool,
    text_seen: &mut bool,
) -> Result<()> {
    match current {
        Element::DocumentContent if root => Ok(()),
        Element::DocumentContent => {
            invalid("OTH content.xml has a nested or duplicate document-content root")
        },
        Element::Body if parent == Some(Element::DocumentContent) && !*body_seen => {
            *body_seen = true;
            Ok(())
        },
        Element::Body => {
            invalid("OTH office:body must occur once directly inside document-content")
        },
        Element::Text if parent == Some(Element::Body) && !*text_seen => {
            *text_seen = true;
            Ok(())
        },
        Element::Text => invalid("OTH office:text must occur once directly inside office:body"),
        Element::Paragraph | Element::Other if root => invalid(
            "OTH content.xml root must be office:document-content in the ODF office namespace",
        ),
        Element::Paragraph | Element::Other => Ok(()),
    }
}

fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> Element {
    let (namespace, local) = reader.resolver().resolve_element(name);
    match namespace {
        ResolveResult::Bound(Namespace(value)) if value == OFFICE_NAMESPACE => match local.as_ref()
        {
            b"document-content" => Element::DocumentContent,
            b"body" => Element::Body,
            b"text" => Element::Text,
            _ => Element::Other,
        },
        ResolveResult::Bound(Namespace(value))
            if value == TEXT_NAMESPACE && local.as_ref() == b"p" =>
        {
            Element::Paragraph
        },
        _ => Element::Other,
    }
}

fn append_text(target: Option<&mut String>, text: &str) -> Result<()> {
    let target = target
        .ok_or_else(|| Error::InvalidFormat("OTH paragraph text state is missing".to_string()))?;
    target
        .try_reserve(text.len())
        .map_err(|source| Error::Allocation {
            resource: "OTH paragraph text",
            source,
        })?;
    target.push_str(&text);
    Ok(())
}

fn reserve<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

fn xml_error(error: quick_xml::Error) -> Error {
    Error::InvalidFormat(format!("invalid OTH content.xml: {error}"))
}

fn invalid<T>(message: &str) -> Result<T> {
    Err(Error::InvalidFormat(message.to_string()))
}

#[cfg(test)]
mod tests {
    use super::{paragraphs, validate};

    #[test]
    fn accepts_a_prefix_aliased_text_web_envelope() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" o:version="1.4"><o:body><o:text><t:p>text</t:p></o:text></o:body></o:document-content>"#;
        assert!(validate(xml).is_ok());
    }

    #[test]
    fn refuses_wrong_or_misplaced_family_containers() {
        let wrong_namespace = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="https://example.test/office"><office:body><office:text/></office:body></office:document-content>"#;
        let duplicate_body = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text/></office:body><office:body><office:text/></office:body></office:document-content>"#;
        assert!(validate(wrong_namespace).is_err());
        assert!(validate(duplicate_body).is_err());
    }

    #[test]
    fn refuses_document_types_and_paragraphs_outside_the_text_body() {
        let dtd = r#"<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE office:document-content [<!ENTITY x "unsafe">]><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text/></office:body></office:document-content>"#;
        let misplaced_paragraph = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><text:p>outside</text:p><office:body><office:text/></office:body></office:document-content>"#;
        assert!(validate(dtd).is_err());
        assert!(validate(misplaced_paragraph).is_err());
    }

    #[test]
    fn projects_paragraph_character_data_without_evaluation() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>plain <text:a xlink:href="https://example.test" xmlns:xlink="http://www.w3.org/1999/xlink">link</text:a></text:p></office:text></office:body></office:document-content>"#;
        let projected = paragraphs(xml).unwrap();
        assert_eq!(projected[0].text(), "plain link");
    }
}
