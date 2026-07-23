//! Explicit ODT page-sequence master-page assignments.

use crate::elements::xml::{OFFICE_NAMESPACE, TEXT_NAMESPACE, is_bound};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

const MAX_DOCUMENT_XML_BYTES: usize = 64 * 1_048_576;
const MAX_XML_DEPTH: usize = 256;
const MAX_PAGES: usize = 1_000_000;
const MAX_MASTER_PAGE_NAME_BYTES: usize = 65_536;
const MAX_TOTAL_MASTER_PAGE_NAME_BYTES: usize = 16 * 1_048_576;

/// Explicit master-page assignments for the pages in an ODT `text:page-sequence`.
///
/// This model preserves only the ordered assignments stored in `content.xml`. It
/// does not calculate page breaks, resolve master-page definitions, or paginate text.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdtPageSequence {
    /// The `text:master-page-name` value for each explicit `text:page`, in order.
    pub master_page_names: Vec<String>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Frame {
    Other,
    OfficeBody,
    OfficeText,
    PageSequence,
    Page,
}

pub(crate) fn parse_page_sequence(xml: &str) -> Result<Option<OdtPageSequence>> {
    if xml.len() > MAX_DOCUMENT_XML_BYTES {
        return Err(invalid(format!(
            "content XML exceeds the {MAX_DOCUMENT_XML_BYTES} byte page-sequence limit"
        )));
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut sequence = None;
    let mut total_name_bytes = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("XML parsing error: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);

        match event {
            Event::Start(ref element) => {
                let local_name = element.local_name();
                let is_page_sequence = text_element && local_name.as_ref() == b"page-sequence";
                let is_page = text_element && local_name.as_ref() == b"page";

                if is_page_sequence {
                    if stack.last() != Some(&Frame::OfficeText) {
                        return Err(invalid(
                            "text:page-sequence must be a direct office:text child",
                        ));
                    }
                    if sequence.is_some() {
                        return Err(invalid(
                            "office:text contains more than one text:page-sequence",
                        ));
                    }
                    ensure_no_attributes(element, "text:page-sequence")?;
                    sequence = Some(OdtPageSequence {
                        master_page_names: Vec::new(),
                    });
                    push_frame(&mut stack, Frame::PageSequence)?;
                } else if is_page {
                    if stack.last() != Some(&Frame::PageSequence) {
                        return Err(invalid(
                            "text:page must be a direct text:page-sequence child",
                        ));
                    }
                    let sequence = sequence
                        .as_mut()
                        .ok_or_else(|| invalid("missing active text:page-sequence"))?;
                    append_page(sequence, &mut total_name_bytes, &reader, element)?;
                    push_frame(&mut stack, Frame::Page)?;
                } else {
                    if matches!(stack.last(), Some(Frame::PageSequence | Frame::Page)) {
                        return Err(invalid(
                            "text:page-sequence may contain only empty text:page children",
                        ));
                    }
                    let frame = if office_element && local_name.as_ref() == b"body" {
                        Frame::OfficeBody
                    } else if office_element
                        && local_name.as_ref() == b"text"
                        && stack.last() == Some(&Frame::OfficeBody)
                    {
                        Frame::OfficeText
                    } else {
                        Frame::Other
                    };
                    push_frame(&mut stack, frame)?;
                }
            },
            Event::Empty(ref element) => {
                let local_name = element.local_name();
                let is_page_sequence = text_element && local_name.as_ref() == b"page-sequence";
                let is_page = text_element && local_name.as_ref() == b"page";

                if is_page_sequence {
                    if stack.last() != Some(&Frame::OfficeText) {
                        return Err(invalid(
                            "text:page-sequence must be a direct office:text child",
                        ));
                    }
                    if sequence.is_some() {
                        return Err(invalid(
                            "office:text contains more than one text:page-sequence",
                        ));
                    }
                    ensure_no_attributes(element, "text:page-sequence")?;
                    return Err(invalid(
                        "text:page-sequence must contain at least one text:page",
                    ));
                }
                if is_page {
                    if stack.last() != Some(&Frame::PageSequence) {
                        return Err(invalid(
                            "text:page must be a direct text:page-sequence child",
                        ));
                    }
                    let sequence = sequence
                        .as_mut()
                        .ok_or_else(|| invalid("missing active text:page-sequence"))?;
                    append_page(sequence, &mut total_name_bytes, &reader, element)?;
                } else if matches!(stack.last(), Some(Frame::PageSequence | Frame::Page)) {
                    return Err(invalid(
                        "text:page-sequence may contain only empty text:page children",
                    ));
                }
            },
            Event::End(ref element) => {
                let local_name = element.local_name();
                let is_page_sequence = text_element && local_name.as_ref() == b"page-sequence";
                let is_page = text_element && local_name.as_ref() == b"page";
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;

                match frame {
                    Frame::Page if !is_page => {
                        return Err(invalid("unexpected text:page closing element"));
                    },
                    Frame::PageSequence if !is_page_sequence => {
                        return Err(invalid("unexpected text:page-sequence closing element"));
                    },
                    Frame::PageSequence => {
                        if sequence
                            .as_ref()
                            .map_or(true, |value| value.master_page_names.is_empty())
                        {
                            return Err(invalid(
                                "text:page-sequence must contain at least one text:page",
                            ));
                        }
                    },
                    Frame::Other | Frame::OfficeBody | Frame::OfficeText | Frame::Page => {},
                }

                if is_page_sequence && frame != Frame::PageSequence {
                    return Err(invalid("unexpected text:page-sequence closing element"));
                }
                if is_page && frame != Frame::Page {
                    return Err(invalid("unexpected text:page closing element"));
                }
            },
            Event::Text(ref text) => match stack.last() {
                Some(Frame::Page) => return Err(invalid("text:page must be empty")),
                Some(Frame::PageSequence) => {
                    let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        invalid(format!("invalid text:page-sequence whitespace: {error}"))
                    })?;
                    if !value.chars().all(char::is_whitespace) {
                        return Err(invalid("text:page-sequence may not contain character data"));
                    }
                },
                _ => {},
            },
            Event::CData(_) | Event::GeneralRef(_)
                if matches!(stack.last(), Some(Frame::PageSequence | Frame::Page)) =>
            {
                return Err(invalid("text:page-sequence may not contain character data"));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DOCTYPE and processing instructions are not allowed in page sequences",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if !stack.is_empty() {
        return Err(invalid("unterminated content XML"));
    }
    Ok(sequence)
}

fn push_frame(stack: &mut Vec<Frame>, frame: Frame) -> Result<()> {
    if stack.len() >= MAX_XML_DEPTH {
        return Err(invalid(format!(
            "content XML exceeds the {MAX_XML_DEPTH} level page-sequence depth limit"
        )));
    }
    stack.push(frame);
    Ok(())
}

fn ensure_no_attributes(element: &BytesStart<'_>, context: &str) -> Result<()> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid {context} attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        return Err(invalid(format!("{context} must not have attributes")));
    }
    Ok(())
}

fn append_page(
    sequence: &mut OdtPageSequence,
    total_name_bytes: &mut usize,
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<()> {
    if sequence.master_page_names.len() >= MAX_PAGES {
        return Err(invalid(format!(
            "text:page-sequence exceeds the {MAX_PAGES} page limit"
        )));
    }

    let mut master_page_name = None;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid text:page attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if is_bound(&namespace, TEXT_NAMESPACE) && local_name.as_ref() == b"master-page-name" {
            if master_page_name.is_some() {
                return Err(invalid(
                    "text:page has duplicate text:master-page-name attributes",
                ));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| {
                    invalid(format!("invalid text:page master-page-name value: {error}"))
                })?
                .into_owned();
            if value.len() > MAX_MASTER_PAGE_NAME_BYTES {
                return Err(invalid(format!(
                    "text:master-page-name exceeds the {MAX_MASTER_PAGE_NAME_BYTES} byte limit"
                )));
            }
            master_page_name = Some(value);
        } else {
            return Err(invalid(
                "text:page may have only a text:master-page-name attribute",
            ));
        }
    }

    let master_page_name =
        master_page_name.ok_or_else(|| invalid("text:page is missing text:master-page-name"))?;
    *total_name_bytes = total_name_bytes
        .checked_add(master_page_name.len())
        .ok_or_else(|| invalid("text:page-sequence name size overflow"))?;
    if *total_name_bytes > MAX_TOTAL_MASTER_PAGE_NAME_BYTES {
        return Err(invalid(format!(
            "text:page-sequence master-page names exceed the {MAX_TOTAL_MASTER_PAGE_NAME_BYTES} byte limit"
        )));
    }
    sequence.master_page_names.push(master_page_name);
    Ok(())
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn invalid(message: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!("invalid ODT page sequence: {message}"))
}

#[cfg(test)]
mod tests {
    use super::{OdtPageSequence, parse_page_sequence};

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    fn content(body: &str) -> String {
        format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:version="1.3"><o:body><o:text>{body}</o:text></o:body></o:document-content>"#
        )
    }

    #[test]
    fn parses_ordered_master_page_names_with_namespace_aliases() {
        let xml = content(
            r#"<x:page-sequence xmlns:x="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
                <x:page x:master-page-name="First"/>
                <x:page x:master-page-name="Right"/>
                <x:page x:master-page-name="Left"></x:page>
            </x:page-sequence>"#,
        );

        assert_eq!(
            parse_page_sequence(&xml).unwrap(),
            Some(OdtPageSequence {
                master_page_names: vec![
                    "First".to_string(),
                    "Right".to_string(),
                    "Left".to_string(),
                ],
            })
        );
    }

    #[test]
    fn returns_none_for_documents_without_a_page_sequence() {
        let spoofed = content(
            r#"<fake:page-sequence xmlns:fake="urn:example"><fake:page fake:master-page-name="Ignored"/></fake:page-sequence>"#,
        );

        assert_eq!(
            parse_page_sequence(&content("<t:p>Plain text</t:p>")).unwrap(),
            None
        );
        assert_eq!(parse_page_sequence(&spoofed).unwrap(), None);
    }

    #[test]
    fn rejects_invalid_page_sequence_structure() {
        for body in [
            r#"<t:page-sequence/>"#,
            r#"<t:page-sequence><t:page/></t:page-sequence>"#,
            r#"<t:page-sequence><t:page t:master-page-name="A" xml:lang="en"/></t:page-sequence>"#,
            r#"<t:page-sequence><t:page t:master-page-name="A"> </t:page></t:page-sequence>"#,
            r#"<t:page-sequence>not a page<t:page t:master-page-name="A"/></t:page-sequence>"#,
            r#"<t:page-sequence><t:span/><t:page t:master-page-name="A"/></t:page-sequence>"#,
            r#"<t:p><t:page-sequence><t:page t:master-page-name="A"/></t:page-sequence></t:p>"#,
            r#"<t:page t:master-page-name="A"/>"#,
            r#"<t:page-sequence><t:page t:master-page-name="A"/></t:page-sequence><t:page-sequence><t:page t:master-page-name="B"/></t:page-sequence>"#,
        ] {
            assert!(
                parse_page_sequence(&content(body)).is_err(),
                "accepted invalid page sequence: {body}"
            );
        }
    }
}
