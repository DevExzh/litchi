//! Inert ODF note authoring and byte-preserving XML mutations.

use super::{Note, parse_notes};
use crate::elements::xml::{TEXT_NAMESPACE, is_bound};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::ops::Range;

const MAX_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 4_096;

impl Note {
    /// Serialize a validated note as a `text:note` fragment.
    ///
    /// Plain-text body newlines are represented by consecutive `text:p`
    /// elements. Structured bodies retain their validated ODF child nodes.
    /// The fragment declares no namespaces itself; mutation helpers add a
    /// local `text` binding when the selected XML context requires one. All
    /// fields, links, event listeners, scripts, and macro metadata remain
    /// inert while the fragment is serialized.
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(r#"<text:note text:note-class="{}""#, self.class.as_str());
        if let Some(id) = &self.id {
            xml.push_str(&format!(r#" text:id="{}""#, escape_xml(id)));
        }
        xml.push('>');
        xml.push_str("<text:note-citation");
        if let Some(label) = &self.label {
            xml.push_str(&format!(r#" text:label="{}""#, escape_xml(label)));
        }
        xml.push('>');
        xml.push_str(&escape_xml(&self.citation));
        xml.push_str("</text:note-citation>");
        if let Some(rich_body) = &self.rich_body {
            if rich_body.is_empty() {
                xml.push_str("<text:note-body/>");
            } else {
                xml.push_str("<text:note-body>");
                rich_body.write_xml(&mut xml);
                xml.push_str("</text:note-body>");
            }
        } else if self.body.is_empty() {
            xml.push_str("<text:note-body/>");
        } else {
            xml.push_str("<text:note-body>");
            for paragraph in self.body.split('\n') {
                xml.push_str("<text:p>");
                xml.push_str(&escape_xml(paragraph));
                xml.push_str("</text:p>");
            }
            xml.push_str("</text:note-body>");
        }
        xml.push_str("</text:note>");
        Ok(xml)
    }
}

/// Insert a validated note at the end of a `text:p` selected in document order.
///
/// Paragraphs nested in lists, tables, and existing note bodies participate in
/// the ordinal. The helper never evaluates fields, follows links, or executes
/// scripts embedded elsewhere in the ODF package.
pub fn insert_note_xml(xml: &str, paragraph_index: usize, note: &Note) -> Result<String> {
    let scan = validated_scan(xml)?;
    let paragraph = scan.paragraphs.get(paragraph_index).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "note paragraph index {paragraph_index} is out of bounds"
        ))
    })?;
    let fragment = note.to_xml_fragment()?;
    match paragraph {
        ParagraphSite::Paired {
            close_start,
            requires_text_binding,
        } => Ok(format!(
            "{}{}{}",
            &xml[..*close_start],
            bind_fragment_if_needed(fragment, *requires_text_binding),
            &xml[*close_start..]
        )),
        ParagraphSite::Empty {
            start,
            end,
            qname,
            requires_text_binding,
        } => {
            let fragment = bind_fragment_if_needed(fragment, *requires_text_binding);
            expand_empty_paragraph(xml, *start, *end, qname, &fragment)
        },
    }
}

/// Replace one `text:note` selected in document order.
///
/// The replacement writes the public note model, including validated structured
/// body content when the replacement has one.
pub fn replace_note_xml(xml: &str, note_index: usize, replacement: &Note) -> Result<String> {
    let scan = validated_scan(xml)?;
    let site = scan
        .notes
        .get(note_index)
        .ok_or_else(|| Error::InvalidFormat(format!("note index {note_index} is out of bounds")))?;
    let fragment =
        bind_fragment_if_needed(replacement.to_xml_fragment()?, site.requires_text_binding);
    Ok(format!(
        "{}{}{}",
        &xml[..site.span.start],
        fragment,
        &xml[site.span.end..]
    ))
}

/// Remove one `text:note` selected in document order.
pub fn remove_note_xml(xml: &str, note_index: usize) -> Result<String> {
    let scan = validated_scan(xml)?;
    let site = scan
        .notes
        .get(note_index)
        .ok_or_else(|| Error::InvalidFormat(format!("note index {note_index} is out of bounds")))?;
    Ok(format!(
        "{}{}",
        &xml[..site.span.start],
        &xml[site.span.end..]
    ))
}

enum ParagraphSite {
    Paired {
        close_start: usize,
        requires_text_binding: bool,
    },
    Empty {
        start: usize,
        end: usize,
        qname: String,
        requires_text_binding: bool,
    },
}

struct OpenElement {
    local: Vec<u8>,
    paragraph_ordinal: Option<usize>,
    note_ordinal: Option<usize>,
    note_start: Option<usize>,
    paragraph_requires_text_binding: bool,
    note_requires_text_binding: bool,
}

struct NoteSite {
    span: Range<usize>,
    requires_text_binding: bool,
}

struct Scan {
    paragraphs: Vec<ParagraphSite>,
    notes: Vec<NoteSite>,
}

fn validated_scan(xml: &str) -> Result<Scan> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "note mutation XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    let semantic = parse_notes(xml)?;
    let scan = scan_locations(xml)?;
    if semantic.len() != scan.notes.len() {
        return Err(Error::InvalidFormat(
            "note semantic and lexical scans disagree".to_string(),
        ));
    }
    Ok(scan)
}

fn scan_locations(xml: &str) -> Result<Scan> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut open = Vec::<OpenElement>::new();
    let mut paragraphs = Vec::<Option<ParagraphSite>>::new();
    let mut notes = Vec::<Option<NoteSite>>::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid note mutation XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("note mutation XML depth overflow".to_string())
                })?;
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(
                        "note mutation XML nesting exceeds 4096 levels".to_string(),
                    ));
                }
                let local = element.local_name().as_ref().to_vec();
                let paragraph_requires_text_binding =
                    !element.name().as_ref().starts_with(b"text:");
                let paragraph_ordinal = if text_element && local == b"p" {
                    let ordinal = paragraphs.len();
                    paragraphs.push(None);
                    Some(ordinal)
                } else {
                    None
                };
                let note_ordinal = if text_element && local == b"note" {
                    let ordinal = notes.len();
                    notes.push(None);
                    Some(ordinal)
                } else {
                    None
                };
                let note_requires_text_binding = paragraph_requires_text_binding
                    || (note_ordinal.is_some() && declares_text_namespace(element)?);
                let note_start = note_ordinal
                    .is_some()
                    .then(|| event_start(xml, event_end))
                    .transpose()?;
                open.push(OpenElement {
                    local,
                    paragraph_ordinal,
                    note_ordinal,
                    note_start,
                    paragraph_requires_text_binding,
                    note_requires_text_binding,
                });
            },
            Event::Empty(ref element)
                if text_element && element.local_name().as_ref() == b"p" => {
                    let start = event_start(xml, event_end)?;
                    let qname = std::str::from_utf8(element.name().as_ref())
                        .map_err(|_| {
                            Error::InvalidFormat(
                                "note insertion paragraph name is not UTF-8".to_string(),
                            )
                        })?
                        .to_owned();
                    paragraphs.push(Some(ParagraphSite::Empty {
                        start,
                        end: event_end,
                        qname,
                        requires_text_binding: !element.name().as_ref().starts_with(b"text:"),
                    }));
                },
            Event::End(ref element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("note mutation XML stack underflow".to_string())
                })?;
                let open_element = open.pop().ok_or_else(|| {
                    Error::InvalidFormat("note mutation XML stack mismatch".to_string())
                })?;
                if open_element.local.as_slice() != element.local_name().as_ref() {
                    return Err(Error::InvalidFormat(
                        "note mutation XML has mismatched elements".to_string(),
                    ));
                }
                let close_start = event_start(xml, event_end)?;
                if let Some(ordinal) = open_element.paragraph_ordinal {
                    paragraphs[ordinal] = Some(ParagraphSite::Paired {
                        close_start,
                        requires_text_binding: open_element.paragraph_requires_text_binding,
                    });
                }
                if let Some(ordinal) = open_element.note_ordinal {
                    let note_start = open_element.note_start.ok_or_else(|| {
                        Error::InvalidFormat(
                            "note mutation XML has no opening note span".to_string(),
                        )
                    })?;
                    notes[ordinal] = Some(NoteSite {
                        span: note_start..event_end,
                        requires_text_binding: open_element.note_requires_text_binding,
                    });
                }
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not allowed in mutable note XML".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if depth != 0 || !open.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete note mutation XML".to_string(),
        ));
    }
    let paragraphs = paragraphs
        .into_iter()
        .map(|paragraph| {
            paragraph
                .ok_or_else(|| Error::InvalidFormat("incomplete note paragraph scan".to_string()))
        })
        .collect::<Result<Vec<_>>>()?;
    let notes = notes
        .into_iter()
        .map(|note| note.ok_or_else(|| Error::InvalidFormat("incomplete note scan".to_string())))
        .collect::<Result<Vec<_>>>()?;
    Ok(Scan { paragraphs, notes })
}

fn event_start(xml: &str, event_end: usize) -> Result<usize> {
    xml.get(..event_end)
        .and_then(|prefix| prefix.rfind('<'))
        .ok_or_else(|| Error::InvalidFormat("unable to locate note XML event start".to_string()))
}

fn declares_text_namespace(element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid note mutation attribute: {error}"))
        })?;
        if attribute.key.as_ref() == b"xmlns:text" {
            return Ok(true);
        }
    }
    Ok(false)
}

fn expand_empty_paragraph(
    xml: &str,
    start: usize,
    end: usize,
    qname: &str,
    fragment: &str,
) -> Result<String> {
    let raw = &xml[start..end];
    let slash = raw.rfind("/>").ok_or_else(|| {
        Error::InvalidFormat("invalid empty paragraph during note insertion".to_string())
    })?;
    Ok(format!(
        "{}{}>{}</{}>{}",
        &xml[..start],
        &raw[..slash],
        fragment,
        qname,
        &xml[end..]
    ))
}

fn bind_fragment_if_needed(fragment: String, requires_text_binding: bool) -> String {
    if !requires_text_binding {
        return fragment;
    }
    fragment.replacen(
        " text:note-class=",
        &format!(r#" xmlns:text="{TEXT_NAMESPACE_STRING}" text:note-class="#),
        1,
    )
}

const TEXT_NAMESPACE_STRING: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        OdfMetaFieldAttribute, OdfMetaFieldElement, OdfMetaFieldNode, OdfNoteBodyContent,
        odt::NoteClass,
    };

    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    fn document(body: &str) -> String {
        format!(
            r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="{TEXT}">{body}</office:text>"#
        )
    }

    fn rich_body() -> OdfNoteBodyContent {
        OdfNoteBodyContent::new(vec![
            OdfMetaFieldNode::Element(OdfMetaFieldElement {
                namespace_uri: TEXT.to_string(),
                local_name: "p".to_string(),
                attributes: Vec::new(),
                children: vec![
                    OdfMetaFieldNode::Text("First ".to_string()),
                    OdfMetaFieldNode::Element(OdfMetaFieldElement {
                        namespace_uri: TEXT.to_string(),
                        local_name: "span".to_string(),
                        attributes: vec![OdfMetaFieldAttribute {
                            namespace_uri: TEXT.to_string(),
                            local_name: "style-name".to_string(),
                            value: "Emphasis".to_string(),
                        }],
                        children: vec![OdfMetaFieldNode::Text("styled".to_string())],
                    }),
                ],
            }),
            OdfMetaFieldNode::Element(OdfMetaFieldElement {
                namespace_uri: TEXT.to_string(),
                local_name: "p".to_string(),
                attributes: Vec::new(),
                children: vec![OdfMetaFieldNode::Text("Second".to_string())],
            }),
        ])
        .unwrap()
    }

    #[test]
    fn note_serialization_and_xml_crud_round_trip() {
        let mut first = Note::new(NoteClass::Footnote, "1&", "First\nSecond").unwrap();
        first.set_id(Some("note&1".to_string())).unwrap();
        first.set_label(Some("*".to_string())).unwrap();
        let fragment = first.to_xml_fragment().unwrap();
        assert!(fragment.contains(r#"text:id="note&amp;1""#));
        assert!(fragment.contains("<text:p>First</text:p><text:p>Second</text:p>"));

        let original = document("<text:p>Before</text:p><text:p/>");
        let inserted = insert_note_xml(&original, 1, &first).unwrap();
        assert_eq!(parse_notes(&inserted).unwrap(), vec![first.clone()]);
        assert!(inserted.contains("<text:p><text:note"));

        let replacement = Note::new(NoteClass::Endnote, "i", "Replacement").unwrap();
        let replaced = replace_note_xml(&inserted, 0, &replacement).unwrap();
        assert_eq!(parse_notes(&replaced).unwrap(), vec![replacement.clone()]);
        let removed = remove_note_xml(&replaced, 0).unwrap();
        assert!(parse_notes(&removed).unwrap().is_empty());
        assert!(removed.contains("<text:p></text:p>"));
    }

    #[test]
    fn rich_note_authoring_uses_validated_structure_and_libreoffice_fixture() {
        let mut note = Note::with_rich_body(NoteClass::Footnote, "*", rich_body()).unwrap();
        note.set_id(Some("rich-note".to_string())).unwrap();
        note.set_label(Some("custom".to_string())).unwrap();
        assert_eq!(note.body(), "First styled\nSecond");
        assert!(note.rich_body().is_some());

        let fragment = note.to_xml_fragment().unwrap();
        assert!(fragment.contains("<text:note-body><text:p xmlns:text="));
        assert!(fragment.contains(r#"text:style-name="Emphasis""#));
        let semantic = parse_notes(&document(&fragment)).unwrap();
        assert_eq!(semantic, vec![note.clone()]);

        let fixture = include_str!(
            "../../../../../test-data/libreoffice-core/writerperfect/qa/unit/data/writer/epubexport/footnote.fodt"
        );
        let inserted = insert_note_xml(fixture, 0, &note).unwrap();
        let inserted_notes = parse_notes(&inserted).unwrap();
        assert_eq!(inserted_notes.len(), 2);
        assert_eq!(inserted_notes[0].body(), "Footnote content");
        assert_eq!(inserted_notes[1].body(), "First styled\nSecond");
        assert!(inserted.contains(r#"text:style-name="Emphasis""#));

        let replacement = Note::with_rich_body(NoteClass::Endnote, "i", rich_body()).unwrap();
        let replaced = replace_note_xml(&inserted, 0, &replacement).unwrap();
        let replaced_notes = parse_notes(&replaced).unwrap();
        assert_eq!(replaced_notes[0].class(), NoteClass::Endnote);
        assert_eq!(replaced_notes[0].body(), "First styled\nSecond");

        note.set_body("Plain replacement").unwrap();
        assert!(note.rich_body().is_none());
        assert!(
            note.to_xml_fragment()
                .unwrap()
                .contains("<text:p>Plain replacement</text:p>")
        );
    }

    #[test]
    fn note_mutations_use_document_order_and_bind_text_when_needed() {
        let nested = document(
            r#"<text:p>A<text:note text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>B<text:note text:note-class="endnote"><text:note-citation>i</text:note-citation><text:note-body/></text:note></text:p></text:note-body></text:note></text:p>"#,
        );
        let outer = Note::new(NoteClass::Footnote, "2", "Outer").unwrap();
        let replaced = replace_note_xml(&nested, 0, &outer).unwrap();
        assert_eq!(parse_notes(&replaced).unwrap(), vec![outer]);

        let alternate_prefix = r#"<root xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:p/><other xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"/></root>"#;
        let inserted = insert_note_xml(
            alternate_prefix,
            0,
            &Note::new(NoteClass::Footnote, "1", "Body").unwrap(),
        )
        .unwrap();
        assert!(
            inserted.contains(
                r#"<text:note xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0""#
            )
        );

        let self_bound_note = r#"<root><text:note xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body/></text:note></root>"#;
        let replaced = replace_note_xml(
            self_bound_note,
            0,
            &Note::new(NoteClass::Endnote, "i", "Replacement").unwrap(),
        )
        .unwrap();
        assert!(
            replaced.contains(
                r#"<text:note xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0""#
            )
        );
        assert_eq!(
            parse_notes(&replaced).unwrap(),
            vec![Note::new(NoteClass::Endnote, "i", "Replacement").unwrap()]
        );
    }

    #[test]
    fn note_authoring_rejects_invalid_text_and_out_of_bounds_locations() {
        assert!(Note::new(NoteClass::Footnote, "\0", "Body").is_err());
        let xml = document("<text:p>Only</text:p>");
        let note = Note::new(NoteClass::Footnote, "1", "Body").unwrap();
        assert!(insert_note_xml(&xml, 1, &note).is_err());
        assert!(replace_note_xml(&xml, 0, &note).is_err());
        assert!(remove_note_xml(&xml, 0).is_err());
    }
}
