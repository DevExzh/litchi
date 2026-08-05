//! Explicit ODT page-sequence master-page assignments.

use crate::elements::xml::{OFFICE_NAMESPACE, TEXT_NAMESPACE, is_bound};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::ops::Range;

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
pub struct Sequence {
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

pub(crate) fn parse_page_sequence(xml: &str) -> Result<Option<Sequence>> {
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
                    sequence = Some(Sequence {
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
                            .is_none_or(|value| value.master_page_names.is_empty())
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
    sequence: &mut Sequence,
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

impl Sequence {
    /// Create a validated explicit page sequence (ODF 1.3 §5.3
    /// `text:page-sequence` with §5.4 `text:page` children).
    ///
    /// Master-page names are stored lexically and never resolved against
    /// `styles.xml`; at least one page is required, matching the schema.
    pub fn new(master_page_names: Vec<String>) -> Result<Self> {
        let sequence = Self { master_page_names };
        validate_sequence(&sequence)?;
        Ok(sequence)
    }

    /// Serialize the sequence in canonical `text:page-sequence` form.
    ///
    /// The fragment is self-checked by re-parsing it, so an authored value can
    /// never serialize into something the reader would reject.
    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_sequence(self)?;
        let mut fragment = format!(
            "<text:page-sequence xmlns:text=\"{}\">",
            String::from_utf8_lossy(TEXT_NAMESPACE)
        );
        for name in &self.master_page_names {
            fragment.push_str("<text:page text:master-page-name=\"");
            fragment.push_str(&litchi_core::xml::escape_xml(name));
            fragment.push_str("\"/>");
        }
        fragment.push_str("</text:page-sequence>");
        let wrapper = format!(
            "<office:document-content xmlns:office=\"{}\"><office:body><office:text>{fragment}</office:text></office:body></office:document-content>",
            String::from_utf8_lossy(OFFICE_NAMESPACE),
        );
        if parse_page_sequence(&wrapper)?.as_ref() != Some(self) {
            return Err(invalid("canonical page-sequence fragment did not validate"));
        }
        Ok(fragment)
    }
}

fn validate_sequence(sequence: &Sequence) -> Result<()> {
    if sequence.master_page_names.is_empty() {
        return Err(invalid(
            "text:page-sequence requires at least one text:page",
        ));
    }
    if sequence.master_page_names.len() > MAX_PAGES {
        return Err(invalid(format!(
            "text:page-sequence exceeds the {MAX_PAGES} page limit"
        )));
    }
    let mut total_name_bytes = 0usize;
    for name in &sequence.master_page_names {
        if name.is_empty() {
            return Err(invalid("text:master-page-name must not be empty"));
        }
        if name
            .bytes()
            .any(|byte| matches!(byte, b'\t' | b'\n' | b'\r'))
        {
            return Err(invalid(
                "text:master-page-name must not contain attribute-normalized whitespace",
            ));
        }
        if name.len() > MAX_MASTER_PAGE_NAME_BYTES {
            return Err(invalid(format!(
                "text:master-page-name exceeds the {MAX_MASTER_PAGE_NAME_BYTES} byte limit"
            )));
        }
        total_name_bytes = total_name_bytes
            .checked_add(name.len())
            .ok_or_else(|| invalid("text:page-sequence name size overflow"))?;
        if total_name_bytes > MAX_TOTAL_MASTER_PAGE_NAME_BYTES {
            return Err(invalid(format!(
                "text:page-sequence master-page names exceed the {MAX_TOTAL_MASTER_PAGE_NAME_BYTES} byte limit"
            )));
        }
    }
    Ok(())
}

/// Where a `text:page-sequence` lives or must be inserted in `content.xml`.
enum PageSequenceSite {
    /// Byte range of the existing `text:page-sequence` element.
    Existing(Range<usize>),
    /// Offset right after the `office:text` start tag; the sequence belongs
    /// there because the schema orders it ahead of all other content.
    FirstChild { content_start: usize },
    /// An empty `<office:text/>` element that must be expanded.
    EmptyText {
        start: usize,
        end: usize,
        qualified_name: String,
    },
}

/// Locate the existing sequence or the insertion site.
///
/// The document was already validated by [`parse_page_sequence`], so at most
/// one sequence exists and `office:text` is well-formed.
fn locate_site(xml: &str) -> Result<PageSequenceSite> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut text_content_start = None;
    let mut empty_text = None;

    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("XML parsing error: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);
        let event = event.into_owned();
        let end_of_event = reader.buffer_position() as usize;

        match event {
            Event::Start(ref element) => {
                let local_name = element.local_name();
                if text_element
                    && local_name.as_ref() == b"page-sequence"
                    && stack.last() == Some(&Frame::OfficeText)
                {
                    // Find the matching end; validated documents have no
                    // nested page-sequence elements, but `text:page` children
                    // may appear as start/end pairs.
                    let mut depth = 0usize;
                    loop {
                        let (_, nested) = reader
                            .read_resolved_event_into(&mut buffer)
                            .map_err(|error| invalid(format!("XML parsing error: {error}")))?;
                        let nested = nested.into_owned();
                        match nested {
                            Event::Start(_) => depth += 1,
                            Event::End(_) => {
                                if depth == 0 {
                                    return Ok(PageSequenceSite::Existing(
                                        event_start..reader.buffer_position() as usize,
                                    ));
                                }
                                depth -= 1;
                            },
                            Event::Eof => {
                                return Err(invalid("unterminated text:page-sequence"));
                            },
                            _ => {},
                        }
                        buffer.clear();
                    }
                }
                let frame = if office_element && local_name.as_ref() == b"body" {
                    Frame::OfficeBody
                } else if office_element
                    && local_name.as_ref() == b"text"
                    && stack.last() == Some(&Frame::OfficeBody)
                {
                    text_content_start = Some(end_of_event);
                    Frame::OfficeText
                } else {
                    Frame::Other
                };
                push_frame(&mut stack, frame)?;
            },
            Event::Empty(ref element) => {
                let local_name = element.local_name();
                if office_element
                    && local_name.as_ref() == b"text"
                    && stack.last() == Some(&Frame::OfficeBody)
                {
                    let qualified_name =
                        String::from_utf8_lossy(element.name().as_ref()).into_owned();
                    empty_text = Some((event_start, end_of_event, qualified_name));
                }
            },
            Event::End(_) => {
                stack.pop();
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if let Some((start, end, qualified_name)) = empty_text {
        return Ok(PageSequenceSite::EmptyText {
            start,
            end,
            qualified_name,
        });
    }
    let content_start = text_content_start
        .ok_or_else(|| invalid("content XML is missing the office:text element"))?;
    Ok(PageSequenceSite::FirstChild { content_start })
}

/// Insert, replace, or remove the `text:page-sequence` of a content document.
///
/// `None` removes an existing sequence and is a no-op when none exists. A new
/// sequence is inserted as the first child of `office:text`, matching the
/// element order of ODF 1.3 §5.1 (`office:text`) and §5.3.
pub(crate) fn set_page_sequence_xml(xml: &str, sequence: Option<&Sequence>) -> Result<String> {
    parse_page_sequence(xml)?;
    let site = locate_site(xml)?;
    let fragment = sequence.map(Sequence::to_xml_fragment).transpose()?;
    let updated = match (site, fragment) {
        (PageSequenceSite::Existing(range), replacement) => super::header_footer::replace_range(
            xml,
            range.start,
            range.end,
            replacement.as_deref().unwrap_or(""),
        )?,
        (PageSequenceSite::FirstChild { content_start }, Some(fragment)) => {
            super::header_footer::replace_range(xml, content_start, content_start, &fragment)?
        },
        (
            PageSequenceSite::EmptyText {
                start,
                end,
                qualified_name,
            },
            Some(fragment),
        ) => {
            let empty = &xml[start..end];
            let marker = empty
                .rfind("/>")
                .ok_or_else(|| invalid("malformed empty office:text element"))?;
            let mut expanded = String::with_capacity(empty.len() + fragment.len() + 16);
            expanded.push_str(&empty[..marker]);
            expanded.push('>');
            expanded.push_str(&fragment);
            expanded.push_str("</");
            expanded.push_str(&qualified_name);
            expanded.push('>');
            super::header_footer::replace_range(xml, start, end, &expanded)?
        },
        (_, None) => return Ok(xml.to_string()),
    };
    // The rewritten document must read back exactly what was authored.
    if parse_page_sequence(&updated)? != sequence.cloned() {
        return Err(invalid("rewritten page sequence did not validate"));
    }
    Ok(updated)
}

fn invalid(message: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!("invalid ODT page sequence: {message}"))
}

#[cfg(test)]
mod tests {
    use super::{Sequence, parse_page_sequence};

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
            Some(Sequence {
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

    #[test]
    fn authors_validates_and_serializes_sequences() {
        use super::set_page_sequence_xml;

        let sequence =
            Sequence::new(vec!["First".to_string(), "Left & Right".to_string()]).unwrap();
        let fragment = sequence.to_xml_fragment().unwrap();
        assert_eq!(
            fragment,
            r#"<text:page-sequence xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><text:page text:master-page-name="First"/><text:page text:master-page-name="Left &amp; Right"/></text:page-sequence>"#
        );

        assert!(Sequence::new(Vec::new()).is_err());
        assert!(Sequence::new(vec![String::new()]).is_err());
        assert!(Sequence::new(vec!["Bad\nName".to_string()]).is_err());

        // Insert as the first child of office:text, ahead of other content.
        let xml = content("<t:p>text</t:p>");
        let updated = set_page_sequence_xml(&xml, Some(&sequence)).unwrap();
        let sequence_start = updated.find("<text:page-sequence").unwrap();
        let paragraph_start = updated.find("<t:p>").unwrap();
        assert!(sequence_start < paragraph_start);
        assert_eq!(
            parse_page_sequence(&updated).unwrap(),
            Some(sequence.clone())
        );

        // Replace an existing sequence, then remove it.
        let replacement = Sequence::new(vec!["Only".to_string()]).unwrap();
        let updated = set_page_sequence_xml(&updated, Some(&replacement)).unwrap();
        assert_eq!(parse_page_sequence(&updated).unwrap(), Some(replacement));
        let updated = set_page_sequence_xml(&updated, None).unwrap();
        assert_eq!(parse_page_sequence(&updated).unwrap(), None);
        assert!(!updated.contains("page-sequence"));
        // Removing again is a no-op.
        assert_eq!(set_page_sequence_xml(&updated, None).unwrap(), updated);

        // An empty office:text element is expanded.
        let empty = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:version="1.3"><o:body><o:text/></o:body></o:document-content>"#
        );
        let updated = set_page_sequence_xml(&empty, Some(&sequence)).unwrap();
        assert_eq!(parse_page_sequence(&updated).unwrap(), Some(sequence));
        assert!(updated.contains("</o:text>"));
    }
}
